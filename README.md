# Student Grade Portal

A Flask web app for calculating student grades from an uploaded Excel workbook.
It provides class statistics, charts, printable report cards, and Excel/CSV downloads.

## Project layout

- `app.py`: Flask web application
- `student_grade_calculator.py`: grade calculation logic
- `sample_students.xlsx`: sample workbook for the demo button
- `templates/`: HTML templates
- `static/style.css`: styling
- `requirements.txt`: Python dependencies
- `render.yaml`: Render deployment configuration

## Deploy online using Render

1. Create or sign in to a GitHub account.
2. Create a new repository and upload the contents of this folder.
3. Sign in to Render and choose **New +** → **Web Service**.
4. Connect the GitHub repository.
5. Render can read `render.yaml`; if entering commands manually, use:
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn app:app`
6. Deploy. Render will provide a public `*.onrender.com` URL that does not contain `claude.ai`.

## Run locally

```bash
pip install -r requirements.txt
python app.py
```

Then open `http://127.0.0.1:5000`.

## Excel format

Required columns: `RollNo`, `Name`, `English`, `Maths`, `Science`, `Computer`, `Hindi`.
Title rows above the header row are supported.

## Important limitations

- Analyses are stored in server memory and disappear when the service restarts.
- This starter version is not designed for confidential student records or multi-user production use.
- Use fictional/demo data for public demos. Do not upload sensitive student data.
- Set a strong `SECRET_KEY` environment variable in the hosting dashboard for deployment.
