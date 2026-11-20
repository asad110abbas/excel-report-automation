# 📑 BRIHA · Excel Report Automation

![BRIHA](https://img.shields.io/badge/BRIHA-Genesis_of_AGI-8B5CF6?style=for-the-badge&labelColor=0B1024) ![Module](https://img.shields.io/badge/Briha-Data_Lab-00F0FF?style=for-the-badge&labelColor=0B1024)
![Python](https://img.shields.io/badge/Python-3776AB?logo=python&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-217346?logo=microsoftexcel&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-charts_%26_formatting-1F6FEB)
![CI](https://github.com/asad110abbas/excel-report-automation/actions/workflows/ci.yml/badge.svg)

> ⚡ **BRIHA: The Genesis of AGI.** This is a **Briha Data Lab** module, part of the [BRIHA AI World](https://github.com/asad110abbas), built by **Syed Asad Abbas**.

Replaces a manual, hours-long monthly Excel routine with **one command**. It reads raw sales rows (CSV or XLSX) and
produces a polished, multi-sheet management report with KPIs, **native Excel charts**, **conditional
formatting** against targets and live `SUM` formulas.

## What the report contains
| Sheet | Content |
|---|---|
| **Summary** | Title, period, KPI block (revenue, units, target, achievement %, best branch) and a revenue-by-category bar chart |
| **By Branch** | Branch table with red/green achievement highlighting, color-scaled revenue and a totals row with formulas |
| **Daily Trend** | Daily revenue table and a line chart |
| **Raw Data** | All source rows with frozen header and auto-filter |

## Run
```bash
pip install -r requirements.txt
python -m reporter.cli                                    # demo with synthetic data
python -m reporter.cli --input my_sales.csv --out reports/nov.xlsx --title "November Sales"
pytest -q
```
Input columns: `date, branch, category, units, revenue, target`.

## Why it matters
- Turns a recurring manual task into a reproducible, version-controlled process
- Output stays **native Excel** (real charts and formulas), so business users keep working the way they already do
- Easy to schedule (Task Scheduler / cron) and email automatically

> The demo data is synthetic.
