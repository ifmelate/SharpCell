# Builds the uncalculated input of tests/corpus/excel-web/hidden-rows.xlsx: a row hidden by hand in ranges, tables and an autoFilter without criteria.
import sys, zipfile
from xml.sax.saxutils import escape

out = sys.argv[1]
strings = []
def sst(s):
    if s not in strings: strings.append(s)
    return strings.index(s)

rows = {}          # row -> {col: xml}
def put(ref, xml):
    col = ''.join(c for c in ref if c.isalpha()); row = int(ref[len(col):])
    rows.setdefault(row, {})[col] = xml
def num(ref, v): put(ref, f'<c r="{ref}"><v>{v}</v></c>')
def text(ref, s): put(ref, f'<c r="{ref}" t="s"><v>{sst(s)}</v></c>')
def formula(ref, f): put(ref, f'<c r="{ref}"><f>{escape(f)}</f></c>')

# Each block holds 1, 2, 4, 8, 16 with the third value's row hidden by hand: sum 31, without it 27.
blocks = [('A', 'plain range', 2, None), ('B', 'sheet autoFilter, no criteria', 10, 'H'),
          ('C', 'table without header row', 20, None), ('D', 'table with header and autoFilter', 30, 'Num'),
          ('E', 'table with header, no autoFilter', 40, 'Num')]
hidden = set()
for key, label, top, header in blocks:
    first = top + 1 if header else top
    if header: text(f'A{top}', header)
    for r, v in zip(range(first, first + 5), [1, 2, 4, 8, 16]): num(f'A{r}', v)
    hidden.add(first + 2)
    text(f'B{top}', f'Block {key}: {label}')
    rng = f'A{first}:A{first + 4}'
    for i, f in enumerate([f'SUBTOTAL(9,{rng})', f'SUBTOTAL(109,{rng})', f'_xlfn.AGGREGATE(9,4,{rng})', f'_xlfn.AGGREGATE(9,5,{rng})']):
        row = 50 + 5 * [b[0] for b in blocks].index(key) + i
        text(f'C{row}', f'Block {key}: ' + f.replace('_xlfn.', '')); formula(f'D{row}', f)

HID = ' hidden="1"'
def colkey(c): return (len(c), c)
sheet_rows = ''.join(
    f'<row r="{r}"{HID if r in hidden else ""}>' + ''.join(rows[r][c] for c in sorted(rows[r], key=colkey)) + '</row>'
    for r in sorted(rows))
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
sheet = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{NS}" xmlns:r="{RNS}">'
         f'<sheetData>{sheet_rows}</sheetData><autoFilter ref="A10:A15"/><tableParts count="3"><tablePart r:id="rId1"/><tablePart r:id="rId2"/><tablePart r:id="rId3"/></tableParts></worksheet>')
style_info = '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
def table(i, name, ref, cols, extra):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="{NS}" id="{i}" name="{name}" displayName="{name}" ref="{ref}" {extra}>'
            + cols + f'{style_info}</table>')
one = '<tableColumns count="1"><tableColumn id="1" name="Num"/></tableColumns>'
table1 = table(1, 'NoHead', 'A20:A24', '<tableColumns count="1"><tableColumn id="1" name="Column1"/></tableColumns>', 'headerRowCount="0" totalsRowShown="0"')
table2 = table(2, 'WithHead', 'A30:A35', '<autoFilter ref="A30:A35"/>' + one, 'totalsRowShown="0"')
table3 = table(3, 'NoButton', 'A40:A45', one, 'totalsRowShown="0"')
sst_xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{NS}" count="{len(strings)}" uniqueCount="{len(strings)}">'
           + ''.join(f'<si><t xml:space="preserve">{escape(s)}</t></si>' for s in strings) + '</sst>')
styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="{NS}">'
          '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
workbook = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{NS}" xmlns:r="{RNS}">'
            '<sheets><sheet name="Spike" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Spike!$A$10:$A$15</definedName></definedNames><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
def rels(items): return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PR}">'
                         + ''.join(f'<Relationship Id="{i}" Type="{T}{t}" Target="{g}"/>' for i, t, g in items) + '</Relationships>')
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'
content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    f'<Override PartName="/xl/workbook.xml" ContentType="{CT}sheet.main+xml"/>'
    f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{CT}worksheet+xml"/>'
    f'<Override PartName="/xl/tables/table1.xml" ContentType="{CT}table+xml"/>'
    f'<Override PartName="/xl/tables/table2.xml" ContentType="{CT}table+xml"/><Override PartName="/xl/tables/table3.xml" ContentType="{CT}table+xml"/>'
    f'<Override PartName="/xl/sharedStrings.xml" ContentType="{CT}sharedStrings+xml"/>'
    f'<Override PartName="/xl/styles.xml" ContentType="{CT}styles+xml"/></Types>')

with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', rels([('rId1', 'officeDocument', 'xl/workbook.xml')]))
    z.writestr('xl/workbook.xml', workbook)
    z.writestr('xl/_rels/workbook.xml.rels', rels([('rId1', 'worksheet', 'worksheets/sheet1.xml'),
        ('rId2', 'sharedStrings', 'sharedStrings.xml'), ('rId3', 'styles', 'styles.xml')]))
    z.writestr('xl/worksheets/sheet1.xml', sheet)
    z.writestr('xl/worksheets/_rels/sheet1.xml.rels', rels([('rId1', 'table', '../tables/table1.xml'), ('rId2', 'table', '../tables/table2.xml'), ('rId3', 'table', '../tables/table3.xml')]))
    z.writestr('xl/tables/table1.xml', table1)
    z.writestr('xl/tables/table2.xml', table2)
    z.writestr('xl/tables/table3.xml', table3)
    z.writestr('xl/sharedStrings.xml', sst_xml)
    z.writestr('xl/styles.xml', styles)
print(out)
