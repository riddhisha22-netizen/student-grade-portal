STUDENT GRADE PORTAL - PRESENTATION UPDATE

This package contains:
- templates/results.html
- static/style.css

How to use:
1. In your GitHub repository, open templates/results.html and replace its contents with the supplied results.html.
2. Open static/style.css and replace its contents with the supplied style.css.
3. Commit the two changes. If Render is connected to this GitHub repository, it should deploy automatically.

No new Python packages or dashboard page are required. The existing results page keeps its current Flask route names and template data references.

New presentation features:
- Updated purple, indigo, teal and rose colour palette.
- More polished result cards and section styling.
- Light/dark appearance toggle for the results area.
- Search and Pass/Fail/Grade filters, plus a clear-filters button.
- Clickable table headings to sort columns.
- Existing Excel, CSV and full-results PNG download buttons remain.

Note: The image download uses html2canvas from a CDN, so the browser needs an internet connection when you click Download Full Results Image.
