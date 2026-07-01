import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

TIER_FILL = {
    "Missing": PatternFill("solid", fgColor="F4CCCC"),
    "Weak":    PatternFill("solid", fgColor="FCE5CD"),
    "Near":    PatternFill("solid", fgColor="FFF2CC"),
    "Winning": PatternFill("solid", fgColor="D9EAD3"),
}
HEADER_FILL = PatternFill("solid", fgColor="1155CC")
HEADER_FONT = Font(color="FFFFFF", bold=True)

def write_sheet(wb, title, columns, rows, tier_col=None):
    ws = wb.create_sheet(title=title[:31])
    ws.append(columns)
    for c in ws[1]:
        c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = Alignment(vertical="center")
    tier_idx = columns.index(tier_col) + 1 if tier_col and tier_col in columns else None
    for r in rows:
        ws.append([r.get(col, "") for col in columns])
        if tier_idx:
            cell = ws.cell(row=ws.max_row, column=tier_idx)
            fill = TIER_FILL.get(cell.value)
            if fill:
                cell.fill = fill
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(columns, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = min(max(len(str(col)) + 2, 12), 60)
    return ws

def write_workbook(out_path, payload):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("README")
    for line in payload["readme_lines"]:
        ws.append([line])
    write_sheet(wb, "Health — Master Matrix", payload["health_matrix"]["columns"], payload["health_matrix"]["rows"], tier_col="Gap Status")
    write_sheet(wb, "Travel — Master Matrix", payload["travel_matrix"]["columns"], payload["travel_matrix"]["rows"], tier_col="Gap Status")
    write_sheet(wb, "Health — Gaps", payload["health_gaps"]["columns"], payload["health_gaps"]["rows"], tier_col="Gap Status")
    write_sheet(wb, "Travel — Gaps", payload["travel_gaps"]["columns"], payload["travel_gaps"]["rows"], tier_col="Gap Status")
    write_sheet(wb, "Health — URL Map", payload["health_urlmap"]["columns"], payload["health_urlmap"]["rows"])
    write_sheet(wb, "Travel — URL Map", payload["travel_urlmap"]["columns"], payload["travel_urlmap"]["rows"])
    write_sheet(wb, "Review", payload["review"]["columns"], payload["review"]["rows"])
    wb.save(out_path)
