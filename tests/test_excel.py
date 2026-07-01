# tests/test_excel.py
import openpyxl
from keyword_bible import excel

def test_write_workbook_has_eight_sheets(tmp_path):
    block = {"columns": ["Keyword", "Gap Status"], "rows": [{"Keyword": "k", "Gap Status": "Missing"}]}
    payload = {
        "readme_lines": ["Snapshot: 2026-06-29"],
        "health_matrix": block, "travel_matrix": block,
        "health_gaps": block, "travel_gaps": block,
        "health_urlmap": {"columns": ["PB URL"], "rows": [{"PB URL": "u"}]},
        "travel_urlmap": {"columns": ["PB URL"], "rows": [{"PB URL": "u"}]},
        "review": {"columns": ["Keyword", "Reason"], "rows": [{"Keyword": "x", "Reason": "no-signal"}]},
    }
    out = str(tmp_path / "b.xlsx")
    excel.write_workbook(out, payload)
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == [
        "README", "Health — Master Matrix", "Travel — Master Matrix",
        "Health — Gaps", "Travel — Gaps", "Health — URL Map", "Travel — URL Map", "Review",
    ]
    ws = wb["Health — Master Matrix"]
    assert ws["A1"].value == "Keyword" and ws["A2"].value == "k"
