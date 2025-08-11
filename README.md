# AndhraPradesh_PythonPivot

Python scripts to generate formatted Excel pivot tables from weekly **Autocomplete** data for Andhra Pradesh reporting.

## Features
- Reads **Autocomplete** sheet from weekly Excel file.
- Filters **Product Appropriateness Result** (`segment3`) to `"Product Not Appropriate"`.
- Groups by **Event Ending Week** (`SnapDate`).
- Calculates:
  - **Valid** → Sum of Volume and % of Volume (from `segment4` == `"Valid"` and `Volume` column).
  - **Total** → Sum of Volume and % (always 100% per week).
- Creates a **Pivot** sheet with merged headers, grand total row, and Excel formulas.

## Example Output

| Event Ending Week | Valid (Sum of Volume) | Valid (% of Volume) | Total (Sum of Volume) | Total (%) |
|-------------------|----------------------|---------------------|-----------------------|-----------|
| 12/13/2024        | 5                    | 83.3%               | 6                     | 100%      |
| 12/20/2024        | 3                    | 75.0%               | 4                     | 100%      |
| **Grand Total**   | 8                    | 80.0%               | 10                    | 100%      |

> **Note:** The above is just sample dummy data for illustration.

![Pivot Screenshot](images/pivot_screenshot.png)

## Requirements
Install dependencies:
```bash
pip install pandas xlsxwriter openpyxl
Usage
Place your weekly Excel file in the project folder.
Update IN_XLSX and OUT_XLSX variables in the script if needed.
Run:
python make_pivot.py
Open the output Excel file and check the new Pivot sheet.
Project Structure
AndhraPradesh_PythonPivot/
│
├── make_pivot.py                  # Main pivot creation script
├── make_pivot_procedural.py       # Procedural version
├── make_pivot_procedural_clean.py # Warning-free version
├── README.md
├── .gitignore
└── (weekly Excel files not pushed to repo)

