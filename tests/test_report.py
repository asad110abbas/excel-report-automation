from openpyxl import load_workbook

from reporter.build import aggregate, build_report
from reporter.sample_data import sample_sales


def test_aggregate_achievement():
    agg = aggregate(sample_sales())
    b = agg["by_branch"]
    assert (b["achievement"] == b["revenue"] / b["target"]).all()
    assert b["revenue"].is_monotonic_decreasing


def test_report_has_expected_sheets_and_charts(tmp_path):
    out = build_report(sample_sales(), tmp_path / "r.xlsx")
    wb = load_workbook(out)
    assert wb.sheetnames == ["Summary", "By Branch", "Daily Trend", "Raw Data"]
    assert len(wb["Summary"]._charts) == 1 and len(wb["Daily Trend"]._charts) == 1
    assert wb["By Branch"]["A1"].value == "branch"
    last = wb["By Branch"].max_row
    assert str(wb["By Branch"][f"C{last}"].value).startswith("=SUM(")
