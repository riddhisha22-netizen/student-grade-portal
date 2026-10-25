"""
===================================================================
 STUDENT GRADE PORTAL  (web front-end for student_grade_calculator.py)
===================================================================
 Upload the class marks Excel file in the browser and get:
   - results table (Total, Percentage, Grade, Result) with search/filter
   - class statistics, charts, grade distribution, top performers
   - a printable report card for every student
   - downloadable Excel / CSV results

 All the calculation logic lives in student_grade_calculator.py
 (keep both files in the same folder). This file only adds the website.

 Run :  python app.py          then open  http://127.0.0.1:5000
===================================================================
"""

import base64
import contextlib
import io
import os
import secrets
import threading
from collections import OrderedDict

import pandas as pd
from flask import (Flask, abort, flash, redirect, render_template, request,
                   send_file, url_for)

import student_grade_calculator as sgc      # your existing calculation code

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_FILE = os.path.join(BASE_DIR, "sample_students.xlsx")
GRADE_ORDER = ["A+", "A", "B+", "B", "C", "D", "F"]
MAX_STORED_RESULTS = 50                     # keep only the latest analyses in memory

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024        # 5 MB upload limit
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(16))

RESULTS = OrderedDict()                     # token -> analysis (in memory)
_LOCK = threading.Lock()                    # sgc uses module-level settings, so run one at a time


class PortalError(Exception):
    """A problem with the uploaded file that we can show to the user."""


# ---------------------- CALCULATION (wraps your script) ----------------------
def analyse(source, pass_pct: float, per_subject: bool) -> dict:
    """Run the whole grade calculation on an Excel file / file-like object."""
    with _LOCK:
        saved = (sgc.PASS_PERCENTAGE, sgc.REQUIRE_PASS_IN_EACH_SUBJECT,
                 sgc.CHART_FILE, sgc.OUTPUT_FILE)
        try:
            sgc.PASS_PERCENTAGE = pass_pct
            sgc.REQUIRE_PASS_IN_EACH_SUBJECT = per_subject

            try:
                df = sgc.load_students(source)
            except ValueError as err:
                # our own "no header row" message mentions RollNo; anything else is a bad file
                if "RollNo" in str(err):
                    raise PortalError(str(err))
                raise PortalError("Could not read that file. Please upload a valid .xlsx Excel file.")
            except Exception:
                raise PortalError("Could not read that file. Please upload a valid .xlsx Excel file.")

            missing = [c for c in ["RollNo", "Name"] + sgc.SUBJECTS if c not in df.columns]
            if missing:
                raise PortalError("Missing required column(s): " + ", ".join(missing) +
                                  ". Expected: RollNo, Name, " + ", ".join(sgc.SUBJECTS))

            df = df[["RollNo", "Name"] + sgc.SUBJECTS].copy()
            if df.empty:
                raise PortalError("No student rows were found under the header row.")

            roll = pd.to_numeric(df["RollNo"], errors="coerce")
            if roll.notna().all():
                df["RollNo"] = roll.astype(int)
            df["Name"] = df["Name"].astype(str).str.strip()

            log = io.StringIO()                                   # capture the script's warnings
            with contextlib.redirect_stdout(log):
                df = sgc.clean_marks(df)
            warnings = [ln.strip() for ln in log.getvalue().splitlines() if ln.strip()]

            df = sgc.process_students(df)
            stats = sgc.class_statistics(df)
            subject_avg = sgc.subject_wise_average(df)

            chart_buf = io.BytesIO()                              # chart in memory, not on disk
            sgc.CHART_FILE = chart_buf
            sgc.plot_performance(df, subject_avg)

            excel_buf = io.BytesIO()                              # Excel in memory, not on disk
            sgc.OUTPUT_FILE = excel_buf
            sgc.export_results(df, stats, subject_avg)
        finally:
            (sgc.PASS_PERCENTAGE, sgc.REQUIRE_PASS_IN_EACH_SUBJECT,
             sgc.CHART_FILE, sgc.OUTPUT_FILE) = saved

    df["Rank"] = df["Percentage"].rank(ascending=False, method="min").astype(int)
    grade_counts = df["Grade"].value_counts().reindex(GRADE_ORDER, fill_value=0)

    return {
        "df": df,
        "stats": stats,
        "subject_avg": subject_avg,
        "grade_counts": grade_counts,
        "top3": df.nlargest(3, "Percentage")[["Name", "Percentage", "Grade"]].to_dict("records"),
        "chart": base64.b64encode(chart_buf.getvalue()).decode(),
        "excel": excel_buf.getvalue(),
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


# ---------------------- ROUTES ----------------------
@app.route("/")
def index():
    return render_template("index.html", subjects=sgc.SUBJECTS,
                           default_pass=sgc.PASS_PERCENTAGE,
                           sample_available=os.path.exists(SAMPLE_FILE))


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
    a = get_analysis(token)
    if a is None:
        return redirect(url_for("index"))
    return render_template("results.html", a=a, token=token, subjects=sgc.SUBJECTS)


@app.route("/report/<token>/<int:idx>")
def report(token, idx):
    a = get_analysis(token)
    if a is None:
        return redirect(url_for("index"))
    df = a["df"]
    if idx < 0 or idx >= len(df):
        abort(404)
    return render_template("report.html", a=a, token=token, s=df.iloc[idx],
                           avg=a["subject_avg"], subjects=sgc.SUBJECTS, total_students=len(df),
                           max_marks=sgc.MAX_MARKS_PER_SUBJECT)


@app.route("/download/<token>/<kind>")
def download(token, kind):
    a = get_analysis(token)
    if a is None:
        return redirect(url_for("index"))
    if kind == "excel":
        return send_file(io.BytesIO(a["excel"]), as_attachment=True,
                         download_name="student_results.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if kind == "csv":
        data = a["df"].drop(columns=["Rank"]).to_csv(index=False).encode("utf-8")
        return send_file(io.BytesIO(data), as_attachment=True,
                         download_name="student_results.csv", mimetype="text/csv")
    if kind == "chart":
        return send_file(io.BytesIO(base64.b64decode(a["chart"])), as_attachment=True,
                         download_name="class_performance.png", mimetype="image/png")
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
