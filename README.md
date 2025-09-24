# SPRINT API Reports Generator

This repository contains a Python application for generating five standardized reports from the commonly updated **SPRINT API dataset** (Excel file).  

The tool uses a Tkinter interface to guide you through selecting the dataset, choosing which reports to run, and monitoring progress via progress bars. Each report is saved as HTML/PNG (and a word cloud for keywords), in the same folder as your dataset.

---

## Features

- **Graphical interface (Tkinter)**
  - Popup dialogs for dataset selection and report choice.
  - Progress bars showing progress within each report.
  - Report-specific prompts to include notes about missing values.

- **Automatic field name cleaning**
  - Strips leading/trailing whitespace from column names.
  - Collapses multiple spaces into single spaces.

- **Column fallback**
  - If expected columns aren’t found, the app prompts you to select from available fields.

- **Five reports available**
  1. **Funding Department Report** – bar chart of projects by funding department.  
  2. **Study Type Report** – pie chart of study type distribution.  
  3. **Public Health Approach Report** – horizontal bar chart of public health approaches.  
  4. **PI Facility Map Report** – geographic scatter plot of PI facilities and locations (geocoded with Geopy).  
  5. **Keyword Analysis Report** – word cloud + top 20 keyword frequency table.  

- **Output formats**
  - HTML and PNG charts (per report).
  - Word cloud image (`wordcloud.png`) for keywords.

---

## Requirements

- **Python 3.8+**
- Dependencies:
  - `pandas`
  - `plotly`
  - `geopy`
  - `wordcloud`
  - `tkinter` (usually comes with Python; may require `python3-tk` on some Linux systems)

Install with:

```bash
pip install pandas plotly geopy wordcloud
```

---

## Usage

1. **Run the script**  
   ```bash
   python sprint_api_reports.py
   ```

2. **Workflow**
   - A popup first asks you to select the SPRINT API dataset (`.xlsx` or `.xls`).
   - A second popup asks which reports to generate:

     ```
     Please select which report(s) to generate:

     1 = Funding Department
     2 = Study Type
     3 = Public Health Approach
     4 = PI Facility Map
     5 = Keyword Analysis
     all = Run All Reports
     ```

   - Progress bars show the steps of each report (cleaning → building chart → saving outputs).
   - At the end, a popup lists which reports were generated.

---

## Output Files

All output is saved in the same folder as the input dataset.

- **Funding Department**
  - `1_SPRINT_API_by_Funder.html`
  - `1_SPRINT_API_by_Funder.png`

- **Study Type**
  - `2_SPRINT_API_Study_Type_Pie_Chart.html`
  - `2_SPRINT_API_Study_Type_Pie_Chart.png`

- **Public Health Approach**
  - `3_SPRINT_API_public_health_approach_chart.html`
  - `3_SPRINT_API_public_health_approach_chart.png`

- **PI Facility Map**
  - `4_SPRINT_API_pi_facility_map.html`
  - `4_SPRINT_API_pi_facility_map.png`

- **Keyword Analysis**
  - `5_SPRINT_API_Keyword_Analysis.html`
  - `wordcloud.png`

---

## Notes

- **Geocoding (PI Facility Map)**
  - Uses OpenStreetMap’s Nominatim service via Geopy.
  - A 1-second rate limit per location is enforced, so this report can take a long time if many unique locations exist.

- **Keyword Analysis**
  - Uses up to four fields: `Key Word 1`, `Key Word 2`, `Key Word 3`, `Key Word 4`.
  - If missing, you will be prompted to select replacement fields.

- **Missing values**
  - Each report asks if you want to include a note about missing values.
