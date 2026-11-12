"""
Student Grade Portal - enhanced version
Keeps the existing Flask portal and grading calculations.
"""
import base64
import contextlib
import io
import os
import secrets
import threading
from collections import OrderedDict

import pandas as pd
from flask import (
    Flask, abort, flash, redirect, render_template, request,
    send_file, url_for
)
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
import matplotlib.pyplot as plt

import student_grade_calculator as sgc

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_FILE = os.path.join(BASE_DIR, "sample_students.xlsx")
GRADE_ORDER = ["A+", "A", "B+", "B", "C", "D", "F"]
MAX_STORED_RESULTS = 50

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(16))

RESULTS = OrderedDict()
_LOCK = threading.Lock()


class PortalError(Exception):
    """A problem with the uploaded file that can be shown to the user."""


def make_grade_chart(grade_counts):
    """Create a grade-distribution chart in memory for the Excel workbook."""
    buffer = io.BytesIO()
    labels = [grade for grade in GRADE_ORDER if int(grade_counts.get(grade, 0)) > 0]
    values = [int(grade_counts.get(grade, 0)) for grade in labels]

    fig, ax = plt.subplots(figsize=(8, 4.5))
    if values and sum(values):
        ax.bar(labels, values, color="#6475d8", width=0.62)
        ax.set_title("Grade Distribution", fontsize=15, fontweight="bold", pad=16)
        ax.set_xlabel("Grade")
        ax.set_ylabel("Number of Students")
        ax.grid(axis="y", alpha=0.2)
        ax.set_axisbelow(True)
        for index, value in enumerate(values):
            ax.text(index, value, str(value), ha="center", va="bottom", fontweight="bold")
    else:
        ax.text(0.5, 0.5, "No grade data available", ha="center", va="center")
        ax.axis("off")

    fig.tight_layout()
    fig.savefig(buffer, format="png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer


def add_charts_to_excel(excel_bytes, performance_chart, grade_counts):
    """Keep the existing workbook sheets and add a Charts sheet with embedded images."""
    workbook_buffer = io.BytesIO(excel_bytes)
    workbook = load_workbook(workbook_buffer)

    if "Charts" in workbook.sheetnames:
        del workbook["Charts"]
    sheet = workbook.create_sheet("Charts")
    sheet["A1"] = "STUDENT PERFORMANCE CHARTS"
    sheet["A1"].font = __import__("openpyxl").styles.Font(
        bold=True, size=16, color="26336B"
    )
    sheet["A2"] = "Class performance and grade distribution"
    sheet.column_dimensions["A"].width = 22

    performance_image = ExcelImage(io.BytesIO(performance_chart))
    performance_image.width = 900
    performance_image.height = 390
    sheet.add_image(performance_image, "A4")

    grade_chart = make_grade_chart(grade_counts)
    grade_image = ExcelImage(grade_chart)
    grade_image.width = 650
    grade_image.height = 360
    sheet.add_image(grade_image, "A25")

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.getvalue()


def analyse(source, pass_pct: float, per_subject: bool) -> dict:
    """Run the original grade calculation and prepare downloads."""
    with _LOCK:
        saved = (
            sgc.PASS_PERCENTAGE,
            sgc.REQUIRE_PASS_IN_EACH_SUBJECT,
            sgc.CHART_FILE,
            sgc.OUTPUT_FILE,
        )
        try:
            sgc.PASS_PERCENTAGE = pass_pct
            sgc.REQUIRE_PASS_IN_EACH_SUBJECT = per_subject

            try:
                df = sgc.load_students(source)
            except ValueError as err:
                if "RollNo" in str(err):
                    raise PortalError(str(err))
                raise PortalError(
                    "Could not read that file. Please upload a valid .xlsx Excel file."
                )
            except Exception:
                raise PortalError(
                    "Could not read that file. Please upload a valid .xlsx Excel file."
                )

            missing = [
                column for column in ["RollNo", "Name"] + sgc.SUBJECTS
                if column not in df.columns
            ]
            if missing:
                raise PortalError(
                    "Missing required column(s): " + ", ".join(missing)
                    + ". Expected: RollNo, Name, " + ", ".join(sgc.SUBJECTS)
                )

            df = df[["RollNo", "Name"] + sgc.SUBJECTS].copy()
            if df.empty:
                raise PortalError("No student rows were found under the header row.")

            roll = pd.to_numeric(df["RollNo"], errors="coerce")
            if roll.notna().all():
                df["RollNo"] = roll.astype(int)
            df["Name"] = df["Name"].astype(str).str.strip()

            log = io.StringIO()
            with contextlib.redirect_stdout(log):
                df = sgc.clean_marks(df)
            warnings = [line.strip() for line in log.getvalue().splitlines() if line.strip()]

            # Original calculation logic remains unchanged.
            df = sgc.process_students(df)
            stats = sgc.class_statistics(df)
            subject_avg = sgc.subject_wise_average(df)

            chart_buffer = io.BytesIO()
            sgc.CHART_FILE = chart_buffer
            sgc.plot_performance(df, subject_avg)

            excel_buffer = io.BytesIO()
            sgc.OUTPUT_FILE = excel_buffer
            sgc.export_results(df, stats, subject_avg)

            chart_bytes = chart_buffer.getvalue()
            excel_bytes = add_charts_to_excel(
                excel_buffer.getvalue(), chart_bytes,
                df["Grade"].value_counts().reindex(GRADE_ORDER, fill_value=0)
            )
        finally:
            (
                sgc.PASS_PERCENTAGE,
                sgc.REQUIRE_PASS_IN_EACH_SUBJECT,
                sgc.CHART_FILE,
                sgc.OUTPUT_FILE,
            ) = saved

    df["Rank"] = df["Percentage"].rank(ascending=False, method="min").astype(int)
    grade_counts = df["Grade"].value_counts().reindex(GRADE_ORDER, fill_value=0)

    return {
        "df": df,
        "stats": stats,
        "subject_avg": subject_avg,
        "grade_counts": grade_counts,
        "top3": df.nlargest(3, "Percentage")[
            ["Name", "Percentage", "Grade"]
        ].to_dict("records"),
        "chart": base64.b64encode(chart_bytes).decode("ascii"),
        "excel": excel_bytes,
        "warnings": warnings,
        "pass_pct": pass_pct,
        "per_subject": per_subject,
    }


def store(analysis: dict) -> str:
    token = secrets.token_urlsafe(8)
    RESULTS[token] = analysis
    while len(RESULTS) > MAX_STORED_RESULTS:
        RESULTS.popitem(last=False)
    return token


def get_analysis(token):
    analysis = RESULTS.get(token)
    if analysis is None:
        flash("That result has expired. Please upload the file again.", "error")
    return analysis


@app.route("/")
def index():
    return render_template(
        "index.html",
        subjects=sgc.SUBJECTS,
        default_pass=sgc.PASS_PERCENTAGE,
        sample_available=os.path.exists(SAMPLE_FILE),
    )


@app.route("/analyze", methods=["POST"])
def analyze():
    try:
        pass_pct = float(request.form.get("pass_pct", sgc.PASS_PERCENTAGE))
        if not 1 <= pass_pct <= 100:
            raise ValueError
    except ValueError:
        flash("Pass percentage must be a number between 1 and 100.", "error")
        return redirect(url_for("index"))

    per_subject = request.form.get("per_subject") == "on"

    try:
        if request.form.get("use_sample"):
            if not os.path.exists(SAMPLE_FILE):
                raise PortalError("Sample file is not available.")
            analysis = analyse(SAMPLE_FILE, pass_pct, per_subject)
            analysis["filename"] = "sample_students.xlsx"
        else:
            upload = request.files.get("file")
            if upload is None or upload.filename == "":
                raise PortalError("Please choose an Excel (.xlsx) file first.")
            if not upload.filename.lower().endswith(".xlsx"):
                raise PortalError("Only .xlsx Excel files are supported.")
            analysis = analyse(io.BytesIO(upload.read()), pass_pct, per_subject)
            analysis["filename"] = os.path.basename(upload.filename)
    except PortalError as err:
        flash(str(err), "error")
        return redirect(url_for("index"))

    return redirect(url_for("results", token=store(analysis)))


@app.route("/results/<token>")
def results(token):
    analysis = get_analysis(token)
    if analysis is None:
        return redirect(url_for("index"))
    return render_template(
        "results.html", a=analysis, token=token, subjects=sgc.SUBJECTS
    )


@app.route("/report/<token>/<int:idx>")
def report(token, idx):
    analysis = get_analysis(token)
    if analysis is None:
        return redirect(url_for("index"))
    df = analysis["df"]
    if idx < 0 or idx >= len(df):
        abort(404)
    return render_template(
        "report.html",
        a=analysis,
        token=token,
        s=df.iloc[idx],
        avg=analysis["subject_avg"],
        subjects=sgc.SUBJECTS,
        total_students=len(df),
        max_marks=sgc.MAX_MARKS_PER_SUBJECT,
    )


@app.route("/download/<token>/<kind>")
def download(token, kind):
    analysis = get_analysis(token)
    if analysis is None:
        return redirect(url_for("index"))

    if kind == "excel":
        return send_file(
            io.BytesIO(analysis["excel"]),
            as_attachment=True,
            download_name="student_results_with_charts.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    if kind == "csv":
        data = analysis["df"].drop(columns=["Rank"]).to_csv(index=False).encode("utf-8")
        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name="student_results.csv",
            mimetype="text/csv",
        )

    if kind == "chart":
        return send_file(
            io.BytesIO(base64.b64decode(analysis["chart"])),
            as_attachment=True,
            download_name="class_performance.png",
            mimetype="image/png",
        )

    abort(404)


@app.errorhandler(413)
def too_large(_):
    flash("File is too large (limit 5 MB).", "error")
    return redirect(url_for("index"))


@app.errorhandler(404)
def not_found(_):
    return render_template("error.html"), 404


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)


