"""
===================================================================
 STUDENT GRADE CALCULATOR - Data Science Mini Project
 BCA 2nd Year
===================================================================
 Reads marks for a whole class from an Excel file
 (Excel_student_grade_calculator__cf6ba8.xlsx),
 calculates Total, Percentage, Grade and Result for each student,
 computes class-level statistics, plots a performance chart, and
 exports everything to a formatted result Excel file.

 Input file  : Excel_student_grade_calculator__cf6ba8.xlsx
               Keep it in the SAME folder as this script. If the exact name is not
               found, the script automatically uses the .xlsx file in that folder.
               Needs columns: RollNo, Name, English, Maths, Science, Computer, Hindi
               The header row can be anywhere (title rows above it are skipped).
               Existing Total/Percentage/Grade/Result columns are ignored and
               recalculated by this script.

 Required libraries : pandas, openpyxl, matplotlib
 Install (if needed): pip install pandas openpyxl matplotlib
===================================================================
"""

import os
import sys
import pandas as pd
import matplotlib
matplotlib.use("Agg")                      # saves charts to file; works without a display
import matplotlib.pyplot as plt
from openpyxl.styles import Font, PatternFill, Alignment

# ---------------------- CONFIGURATION ----------------------
# All files are looked up next to this script, so it works from any folder.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(BASE_DIR, "Excel student grade calculator .xlsx")  # linked input file
OUTPUT_FILE = os.path.join(BASE_DIR, "student_results.xlsx")   # generated results
CHART_FILE = os.path.join(BASE_DIR, "class_performance.png")   # generated chart

SUBJECTS = ["English", "Maths", "Science", "Computer", "Hindi"]
MAX_MARKS_PER_SUBJECT = 100
PASS_PERCENTAGE = 40
REQUIRE_PASS_IN_EACH_SUBJECT = False   # True = student must also score >= 40 in every subject


# ---------------------- CORE LOGIC ----------------------
def calculate_grade(percentage: float) -> str:
    """Return a letter grade for a given percentage."""
    if percentage >= 90:
        return "A+"
    elif percentage >= 80: 
        return "A"
    elif percentage >= 70:
        return "B+"
    elif percentage >= 60:
        return "B"
    elif percentage >= 50:
        return "C"
    elif percentage >= PASS_PERCENTAGE:
        return "D"
    else:
        return "F"


def process_students(df: pd.DataFrame) -> pd.DataFrame:
    """Add Total, Percentage, Grade and Result columns to the dataframe."""
    max_total = len(SUBJECTS) * MAX_MARKS_PER_SUBJECT

    df["Total"] = df[SUBJECTS].sum(axis=1)
    df["Percentage"] = (df["Total"] / max_total * 100).round(2)
    df["Grade"] = df["Percentage"].apply(calculate_grade)

    def result(row):
        passed = row["Percentage"] >= PASS_PERCENTAGE
        if REQUIRE_PASS_IN_EACH_SUBJECT:
            passed = passed and (row[SUBJECTS] >= PASS_PERCENTAGE).all()
        return "Pass" if passed else "Fail"

    df["Result"] = df.apply(result, axis=1)
    return df


def class_statistics(df: pd.DataFrame) -> dict:
    """Compute simple class-level statistics (basic data analysis)."""
    topper_row = df.loc[df["Percentage"].idxmax()]
    return {
        "Number of Students": len(df),
        "Class Average %": round(df["Percentage"].mean(), 2),
        "Highest %": df["Percentage"].max(),
        "Class Topper": topper_row["Name"],
        "Lowest %": df["Percentage"].min(),
        "Pass Count": int((df["Result"] == "Pass").sum()),
        "Fail Count": int((df["Result"] == "Fail").sum()),
        "Pass %": round((df["Result"] == "Pass").mean() * 100, 2),
    }


def subject_wise_average(df: pd.DataFrame) -> pd.Series:
    """Average marks per subject, across the whole class."""
    return df[SUBJECTS].mean().round(2)


# ---------------------- INPUT ----------------------
def find_input_file():
    """
    Locate the marks file.
    1. Use INPUT_FILE if it exists.
    2. Otherwise look in the script's folder for any .xlsx file (ignoring our own
       output file and Excel temp files). Browsers often rename downloads, e.g.
       'Excel_student_grade_calculator_ (1).xlsx', so an exact name can't be trusted.
    3. A path given on the command line always wins:  python script.py "C:/path/marks.xlsx"
    """
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        return sys.argv[1]

    if os.path.exists(INPUT_FILE):
        return INPUT_FILE

    candidates = [
        os.path.join(BASE_DIR, f) for f in os.listdir(BASE_DIR)
        if f.lower().endswith(".xlsx")
        and not f.startswith("~$")
        and f != os.path.basename(OUTPUT_FILE)
    ]
    if not candidates:
        return None

    # Prefer files that look like the student sheet, then the most recently modified
    candidates.sort(key=lambda f: ("student" not in os.path.basename(f).lower()
                                   and "excel" not in os.path.basename(f).lower(),
                                   -os.path.getmtime(f)))
    return candidates[0]


def load_students(path: str) -> pd.DataFrame:
    """Read the Excel file, automatically finding the header row (the one containing 'RollNo')."""
    raw = pd.read_excel(path, header=None)
    header_row = None
    for i in range(len(raw)):
        if "RollNo" in raw.iloc[i].astype(str).str.strip().values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not find a header row containing 'RollNo'.")

    df = raw.iloc[header_row + 1:].copy()
    df.columns = raw.iloc[header_row].astype(str).str.strip()
    df.columns.name = None
    df = df.dropna(subset=["Name"]).reset_index(drop=True)   # skip empty rows
    return df


# ---------------------- VALIDATION ----------------------
def clean_marks(df: pd.DataFrame) -> pd.DataFrame:
    """Make sure marks are numeric and within the allowed range."""
    df[SUBJECTS] = df[SUBJECTS].apply(pd.to_numeric, errors="coerce")

    if df[SUBJECTS].isna().any().any():
        print("WARNING: Some marks are blank or non-numeric. They will be treated as 0.")
        df[SUBJECTS] = df[SUBJECTS].fillna(0)

    out_of_range = (df[SUBJECTS] < 0) | (df[SUBJECTS] > MAX_MARKS_PER_SUBJECT)
    if out_of_range.any().any():
        bad = df.loc[out_of_range.any(axis=1), "Name"].tolist()
        print(f"WARNING: Marks outside 0-{MAX_MARKS_PER_SUBJECT} found for: {bad}. "
              f"They will be clipped to the valid range.")
        df[SUBJECTS] = df[SUBJECTS].clip(0, MAX_MARKS_PER_SUBJECT)
    return df


# ---------------------- VISUALIZATION ----------------------
def plot_performance(df: pd.DataFrame, subject_avg: pd.Series) -> None:
    """Save two charts: overall % per student, and subject-wise class average."""
    fig, axes = plt.subplots(1, 2, figsize=(16, 7))

    # Chart 1: Percentage per student
    colors = ["#4CAF50" if r == "Pass" else "#E53935" for r in df["Result"]]
    axes[0].bar(df["Name"], df["Percentage"], color=colors)
    axes[0].axhline(y=PASS_PERCENTAGE, color="black", linestyle="--",
                    label=f"Pass Mark ({PASS_PERCENTAGE}%)")
    axes[0].set_title("Student-wise Percentage")
    axes[0].set_ylabel("Percentage")
    axes[0].set_ylim(0, 100)
    axes[0].tick_params(axis="x", rotation=45)
    plt.setp(axes[0].get_xticklabels(), ha="right")
    axes[0].legend()

    # Chart 2: Subject-wise class average
    axes[1].bar(subject_avg.index, subject_avg.values, color="#1E88E5")
    axes[1].set_title("Subject-wise Class Average")
    axes[1].set_ylabel("Average Marks")
    axes[1].set_ylim(0, MAX_MARKS_PER_SUBJECT)
    for i, v in enumerate(subject_avg.values):
        axes[1].text(i, v + 1, f"{v:.1f}", ha="center")

    plt.tight_layout()
    plt.savefig(CHART_FILE, dpi=150)
    plt.close()


# ---------------------- EXPORT ----------------------
def export_results(df: pd.DataFrame, stats: dict, subject_avg: pd.Series) -> None:
    """Write results, class summary and subject averages to a formatted Excel file."""
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Results", index=False)
        pd.DataFrame(stats.items(), columns=["Metric", "Value"]).to_excel(
            writer, sheet_name="Class Summary", index=False
        )
        subject_avg.rename("Average Marks").to_excel(
            writer, sheet_name="Subject Averages"
        )

        header_fill = PatternFill("solid", fgColor="1E88E5")
        header_font = Font(bold=True, color="FFFFFF")
        fail_fill = PatternFill("solid", fgColor="FFCDD2")

        for ws in writer.book.worksheets:
            for cell in ws[1]:                       # header row
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
            for col in ws.columns:                   # auto-fit column widths
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = width + 3

        # Highlight failed students in the Results sheet
        ws = writer.book["Results"]
        result_col = list(df.columns).index("Result") + 1
        for row in ws.iter_rows(min_row=2):
            if row[result_col - 1].value == "Fail":
                for cell in row:
                    cell.fill = fail_fill


# ---------------------- MAIN ----------------------
def main():
    global INPUT_FILE
    found = find_input_file()
    if found is None:
        print(f"ERROR: No Excel (.xlsx) marks file found in: {BASE_DIR}")
        print("Put your marks file in the same folder as this script and run it again.")
        print(f"Required columns: RollNo, Name, {', '.join(SUBJECTS)}")
        return
    INPUT_FILE = found
    print(f"Using input file: {INPUT_FILE}")

    try:
        df = load_students(INPUT_FILE)
    except ValueError as err:
        print(f"ERROR: {err}")
        return

    missing = [c for c in ["RollNo", "Name"] + SUBJECTS if c not in df.columns]
    if missing:
        print(f"ERROR: '{INPUT_FILE}' is missing required column(s): {missing}")
        print(f"Expected columns: RollNo, Name, {', '.join(SUBJECTS)}")
        return

    # Keep only the columns we need (drops any old Total/Percentage/Grade/Result)
    df = df[["RollNo", "Name"] + SUBJECTS].copy()

    df = clean_marks(df)
    df = process_students(df)
    stats = class_statistics(df)
    subject_avg = subject_wise_average(df)
    plot_performance(df, subject_avg)
    export_results(df, stats, subject_avg)

    print("\n===== STUDENT RESULTS =====")
    print(df[["RollNo", "Name", "Total", "Percentage", "Grade", "Result"]]
          .to_string(index=False))

    print("\n===== CLASS SUMMARY =====")
    for key, value in stats.items():
        print(f"{key}: {value}")

    print("\n===== SUBJECT-WISE AVERAGE =====")
    print(subject_avg.to_string())

    print(f"\nInput read from      : {INPUT_FILE}")
    print(f"Results exported to  : {OUTPUT_FILE}")
    print(f"Chart saved as       : {CHART_FILE}")


if __name__ == "__main__":
    main()
