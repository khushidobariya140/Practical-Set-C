"""Set C - Training Performance Analysis (Python module)
Run from the repository root:   python python/analysis.py
All paths are resolved relative to the repository root, so no edits are needed.
"""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")                       # works on machines without a display
import matplotlib.pyplot as plt
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs"
OUT.mkdir(exist_ok=True)

# ---------------------------------------------------------------- P1: load, clean, merge
assessments = pd.read_csv(ROOT / "data/raw/assessments.csv",
                          dtype={"course_id": str, "month": str, "batch": str})
courses = pd.read_csv(ROOT / "data/raw/courses.csv", dtype=str)
print(f"Raw assessments rows : {len(assessments)}")

# numeric types
assessments["assessment_id"] = assessments["assessment_id"].astype(int)
for col in ["score", "attendance_pct"]:
    assessments[col] = pd.to_numeric(assessments[col], errors="raise")
assert pd.api.types.is_numeric_dtype(assessments["score"])
assert pd.api.types.is_numeric_dtype(assessments["attendance_pct"])

# remove the exact duplicate row
assessments = assessments.drop_duplicates().reset_index(drop=True)
print(f"After duplicate removal: {len(assessments)}")

# left join on course_id
df = assessments.merge(courses, on="course_id", how="left")
assert len(df) == 12, "Merged data must contain exactly 12 rows"
assert df["department"].notna().all(), "Unmatched course_id found (NaN department)"
print("Assertions passed: 12 rows, zero unmatched course_id")

# ---------------------------------------------------------------- P2: derived field + analysis
df["pass_flag"] = (df["score"] >= 50).astype(int)         # exactly 50 counts as pass

dept = (df.groupby("department")
          .agg(total_assessments=("pass_flag", "count"),
               passing=("pass_flag", "sum"),
               avg_score=("score", "mean"))
          .reset_index())
dept["pass_rate_pct"] = (dept["passing"] / dept["total_assessments"] * 100).round(2)
dept["avg_score"] = dept["avg_score"].round(2)
print("\nDepartment summary:\n", dept.to_string(index=False))

course = (df.groupby(["course_id", "course"])
            .agg(total_assessments=("pass_flag", "count"),
                 passing=("pass_flag", "sum"),
                 avg_score=("score", "mean"))
            .reset_index())
course["pass_rate_pct"] = (course["passing"] / course["total_assessments"] * 100).round(2)
course["avg_score"] = course["avg_score"].round(2)
min_rate = course["pass_rate_pct"].min()
lowest = course[course["pass_rate_pct"] == min_rate]       # reports all ties
print("\nCourse summary:\n", course.to_string(index=False))
for _, r in lowest.iterrows():
    print(f"\nLowest pass rate: {r['course']} ({r['course_id']}) = "
          f"{r['passing']} passing / {r['total_assessments']} total = {r['pass_rate_pct']:.2f}%")

# ---------------------------------------------------------------- P3: chart + exports
month_order = ["Jan", "Feb", "Mar"]
monthly = (df.groupby("month")["score"].mean().reindex(month_order).round(2))
print("\nMonthly average score:\n", monthly.to_string())

fig, ax = plt.subplots(figsize=(7, 4.5))
bars = ax.bar(monthly.index, monthly.values, color="#1F4E79", width=0.55)
ax.bar_label(bars, fmt="%.2f", padding=3, fontsize=10)
ax.set_title("Monthly Average Assessment Score (Jan - Mar)", fontsize=13, fontweight="bold")
ax.set_xlabel("Month")
ax.set_ylabel("Average Score (0-100)")
ax.set_ylim(0, 100)
ax.grid(axis="y", linestyle="--", alpha=0.4)
ax.set_axisbelow(True)
fig.tight_layout()
fig.savefig(OUT / "python_chart.png", dpi=150)
plt.close(fig)

df.to_csv(OUT / "clean_data.csv", index=False)
dept.to_csv(OUT / "python_summary.csv", index=False)
print("\nSaved: outputs/python_chart.png, outputs/clean_data.csv, outputs/python_summary.csv")
