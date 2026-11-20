"""python -m reporter.cli --input sales.csv --out reports/report.xlsx   (omit --input to use sample data)"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from .build import build_report
from .sample_data import sample_sales

REQUIRED = {"date", "branch", "category", "units", "revenue", "target"}


def main():
    ap = argparse.ArgumentParser(description="Turn raw sales rows into a formatted Excel report")
    ap.add_argument("--input", type=Path, help="CSV or XLSX with columns: " + ", ".join(sorted(REQUIRED)))
    ap.add_argument("--out", type=Path, default=Path("reports/monthly_sales_report.xlsx"))
    ap.add_argument("--title", default="Monthly Sales Report")
    a = ap.parse_args()
    if a.input:
        df = pd.read_excel(a.input) if a.input.suffix in {".xlsx", ".xls"} else pd.read_csv(a.input)
        missing = REQUIRED - set(df.columns)
        if missing:
            raise SystemExit(f"input is missing columns: {sorted(missing)}")
        df["date"] = pd.to_datetime(df["date"]).dt.date
    else:
        df = sample_sales()
    print(f"report written to {build_report(df, a.out, a.title)}")


if __name__ == "__main__":
    main()
