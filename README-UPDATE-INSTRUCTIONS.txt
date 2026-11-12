UPDATED FILES — STUDENT GRADE PORTAL

This update does NOT add a dashboard or a new page. It improves the existing
Results page and keeps your original grading module.

Replace these files in your existing GitHub project:

1. app.py
2. templates/results.html
3. static/style.css

Do NOT replace:
- student_grade_calculator.py
- templates/index.html
- templates/base.html
- templates/report.html
- templates/error.html
- requirements.txt
- render.yaml

New features:
- Refreshed styling for the existing results page
- Search, Pass/Fail filter, and Grade filter
- Download Full Results Image button (PNG of the full results section)
- Excel workbook keeps Results, Class Summary, and Subject Averages sheets
- Excel workbook gets an additional Charts sheet with embedded chart images

No new Python package is required if your existing requirements.txt already
includes pandas, openpyxl, matplotlib, and Flask.

After replacing the files, commit the changes to GitHub. Render should redeploy
automatically if it is connected to that repository.

IMPORTANT:
The PNG button uses html2canvas from a CDN, so an internet connection is needed
in the browser when clicking Download Full Results Image.
