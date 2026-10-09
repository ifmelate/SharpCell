"""Writes budget.xlsx, the workbook the sample reads.

The formulas are saved without results, as a script cannot calculate them; the sample
calls Recalculate before reading. Run: python3 make-budget.py (standard library only).
"""
import zipfile
from pathlib import Path

CELLS = {
    "A1": "Item", "B1": "Monthly", "C1": "Yearly", "E1": "Income", "F1": 3000,
    "A2": "Rent", "B2": 1200, "C2": "=B2*12", "E2": "Left", "F2": "=F1-B6",
    "A3": "Food", "B3": 450, "C3": "=B3*12", "E3": "Rent share", "F3": "=B2/F1",
    "A4": "Transport", "B4": 120, "C4": "=B4*12",
    "A5": "Internet", "B5": 40, "C5": "=B5*12",
    "A6": "Total", "B6": "=SUM(B2:B5)", "C6": "=SUM(C2:C5)",
}

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def cell(ref, value):
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"><v>{value}</v></c>'
    if value.startswith("="):
        return f'<c r="{ref}"><f>{value[1:]}</f></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t>{value}</t></is></c>'


def row_of(ref):
    return int("".join(ch for ch in ref if ch.isdigit()))


def column_of(ref):
    return "".join(ch for ch in ref if ch.isalpha())


rows = {}
for ref, value in CELLS.items():
    rows.setdefault(row_of(ref), []).append((column_of(ref), ref, value))
sheet_data = "".join(
    f'<row r="{r}">' + "".join(cell(ref, v) for _, ref, v in sorted(cells)) + "</row>"
    for r, cells in sorted(rows.items())
)

parts = {
    "[Content_Types].xml": (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    ),
    "_rels/.rels": f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    "xl/workbook.xml": f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets><sheet name="Budget" sheetId="1" r:id="rId1"/></sheets></workbook>',
    "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    "xl/worksheets/sheet1.xml": f'<worksheet xmlns="{MAIN}"><sheetData>{sheet_data}</sheetData></worksheet>',
}

target = Path(__file__).with_name("budget.xlsx")
with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
    for name, xml in parts.items():
        info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        archive.writestr(info, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + xml)
print(f"Wrote {target}")
