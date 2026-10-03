# Training Performance Analysis - Set C (Excel | SQL | Python)

**Business question:** Which course needs the most academic support, and how does performance differ across batches?
Pass rule: `pass_flag = 1` when `score >= 50`. Pass rate = passing assessments / total assessments. Duplicate row (assessment 12) removed -> 12 clean rows.

## Folder structure
```
data/raw/        assessments.csv (13 rows incl. duplicate), courses.csv (4 rows)
excel/           analysis.xlsx (Raw, Lookup, Clean, Summary) + build_workbook.py
sql/             setup.sql, queries.sql, run_sql.py  (SQLite 3.35+)
python/          analysis.py
outputs/         clean_data.csv, python_summary.csv, python_chart.png
outputs/sql/     saved result of every query (CSV)
requirements.txt .gitignore
```

## How to run (from the repository root)
```
pip install -r requirements.txt
python python/analysis.py        # writes outputs/clean_data.csv, python_summary.csv, python_chart.png
python sql/run_sql.py            # runs setup.sql then queries.sql, saves outputs/sql/*.csv
# or manually in the sqlite3 shell:  .read sql/setup.sql   then   .read sql/queries.sql
```
Excel: open `excel/analysis.xlsx` (all formulas are live; start at the Summary sheet).

## Key results (identical in Excel, SQL and Python)
| Metric | Value |
|---|---|
| Technology avg score / Business avg score | 56.00 / 67.00 |
| Lowest pass-rate course | Python (C4): 1 passing / 3 total = 33.33% |
| Courses with avg score < 60 | Python 49.33, PowerBI 53.33 |
| Top two batches by avg score | Evening 67.00, Morning 61.25 |
| Passing assessments by batch (Morning / Evening / Weekend) | 3 / 3 / 2 |
