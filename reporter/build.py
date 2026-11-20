"""Build a formatted, multi-sheet Excel management report from raw sales rows.

Sheets: Summary (KPIs + chart), By Branch (conditional formatting vs target), Daily Trend (line chart), Raw Data.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F6FEB")
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="D0D7DE")
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
MONEY = '#,##0'
PCT = '0.0%'


def _write_table(ws, df: pd.DataFrame, start_row: int = 1, formats: dict | None = None) -> None:
    formats = formats or {}
    for j, col in enumerate(df.columns, 1):
        c = ws.cell(start_row, j, col)
        c.fill, c.font, c.alignment, c.border = HEADER_FILL, HEADER_FONT, Alignment(horizontal="center"), BORDER
    for i, row in enumerate(df.itertuples(index=False), start_row + 1):
        for j, v in enumerate(row, 1):
            c = ws.cell(i, j, v)
            c.border = BORDER
            if df.columns[j - 1] in formats:
                c.number_format = formats[df.columns[j - 1]]
    for j, col in enumerate(df.columns, 1):
        width = max(len(str(col)), *(len(f"{v:,.0f}" if isinstance(v, float) else str(v)) for v in df.iloc[:, j - 1]))
        ws.column_dimensions[get_column_letter(j)].width = min(width + 4, 40)
    ws.freeze_panes = ws.cell(start_row + 1, 1)


def aggregate(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    by_branch = df.groupby("branch").agg(units=("units", "sum"), revenue=("revenue", "sum"),
                                         target=("target", "sum")).reset_index()
    by_branch["achievement"] = by_branch["revenue"] / by_branch["target"]
    by_branch = by_branch.sort_values("revenue", ascending=False)
    daily = df.groupby("date", as_index=False)["revenue"].sum()
    by_cat = df.groupby("category", as_index=False)["revenue"].sum().sort_values("revenue", ascending=False)
    return {"by_branch": by_branch, "daily": daily, "by_category": by_cat}


def build_report(df: pd.DataFrame, out: Path, title: str = "Monthly Sales Report") -> Path:
    agg = aggregate(df)
    wb = Workbook()

    # ---- Summary ----
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = title
    ws["A1"].font = Font(size=18, bold=True, color="1B2733")
    ws["A2"] = f"Period: {min(df['date'])} to {max(df['date'])}"
    ws["A2"].font = Font(italic=True, color="56616C")
    total_rev, total_target = df["revenue"].sum(), df["target"].sum()
    kpis = [("Total revenue", total_rev, MONEY), ("Units sold", int(df["units"].sum()), '#,##0'),
            ("Target", total_target, MONEY), ("Achievement", total_rev / total_target, PCT),
            ("Best branch", agg["by_branch"].iloc[0]["branch"], None)]
    for i, (label, value, fmt) in enumerate(kpis, 4):
        ws.cell(i, 1, label).font = Font(bold=True)
        c = ws.cell(i, 2, value)
        if fmt:
            c.number_format = fmt
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 20, 18
    _write_table(ws, agg["by_category"], start_row=11, formats={"revenue": MONEY})
    ws.freeze_panes = None
    chart = BarChart()
    chart.title, chart.y_axis.title = "Revenue by category", "PKR"
    n = len(agg["by_category"])
    chart.add_data(Reference(ws, min_col=2, min_row=11, max_row=11 + n), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=12, max_row=11 + n))
    chart.height, chart.width = 8, 16
    ws.add_chart(chart, "D4")

    # ---- By Branch ----
    wb2 = wb.create_sheet("By Branch")
    b = agg["by_branch"]
    _write_table(wb2, b, formats={"revenue": MONEY, "target": MONEY, "achievement": PCT})
    last = len(b) + 1
    wb2.conditional_formatting.add(f"E2:E{last}", CellIsRule(operator="lessThan", formula=["1"],
                                   fill=PatternFill("solid", fgColor="FDE2E1"), font=Font(color="B42318")))
    wb2.conditional_formatting.add(f"E2:E{last}", CellIsRule(operator="greaterThanOrEqual", formula=["1"],
                                   fill=PatternFill("solid", fgColor="DCFCE7"), font=Font(color="166534")))
    wb2.conditional_formatting.add(f"C2:C{last}", ColorScaleRule(start_type="min", start_color="FFFFFF",
                                                                  end_type="max", end_color="7CC8FF"))
    tot = last + 1
    wb2.cell(tot, 1, "Total").font = Font(bold=True)
    for col in ("B", "C", "D"):
        wb2[f"{col}{tot}"] = f"=SUM({col}2:{col}{last})"
        wb2[f"{col}{tot}"].font = Font(bold=True)
        wb2[f"{col}{tot}"].number_format = MONEY
    wb2[f"E{tot}"] = f"=C{tot}/D{tot}"
    wb2[f"E{tot}"].number_format = PCT

    # ---- Daily Trend ----
    wb3 = wb.create_sheet("Daily Trend")
    _write_table(wb3, agg["daily"], formats={"revenue": MONEY})
    line = LineChart()
    line.title, line.y_axis.title = "Daily revenue", "PKR"
    line.add_data(Reference(wb3, min_col=2, min_row=1, max_row=len(agg["daily"]) + 1), titles_from_data=True)
    line.set_categories(Reference(wb3, min_col=1, min_row=2, max_row=len(agg["daily"]) + 1))
    line.height, line.width = 9, 22
    wb3.add_chart(line, "D2")

    # ---- Raw ----
    wb4 = wb.create_sheet("Raw Data")
    _write_table(wb4, df, formats={"revenue": MONEY, "target": MONEY})
    wb4.auto_filter.ref = wb4.dimensions

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out
