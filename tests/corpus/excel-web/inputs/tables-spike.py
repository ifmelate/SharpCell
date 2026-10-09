# Builds the uncalculated input of tests/corpus/excel-web/tables-spike.xlsx: SUBTOTAL, AGGREGATE and #This Row on a filtered table and a table with a hidden row.
import sys, zipfile
from xml.sax.saxutils import escape

out = sys.argv[1]
strings = []
def sst(s):
    if s not in strings: strings.append(s)
    return strings.index(s)

rows = {}          # row -> {col: xml}
hidden = {3, 5, 6, 14}
def put(ref, xml):
    col = ''.join(c for c in ref if c.isalpha()); row = int(ref[len(col):])
    rows.setdefault(row, {})[col] = xml
def num(ref, v): put(ref, f'<c r="{ref}"><v>{v}</v></c>')
def text(ref, s): put(ref, f'<c r="{ref}" t="s"><v>{sst(s)}</v></c>')
def formula(ref, f): put(ref, f'<c r="{ref}"><f>{escape(f)}</f></c>')

# Table Sales A1:B8: header, six data rows, totals row; filter Grp = x hides rows 3, 5, 6.
text('A1', 'Val'); text('B1', 'Grp')
for r, (v, g) in enumerate([(10, 'x'), (20, 'y'), (30, 'x'), (40, 'y'), (50, 'z'), (60, 'x')], start=2):
    num(f'A{r}', v); text(f'B{r}', g)
formula('A8', 'SUBTOTAL(109,Sales[Val])'); text('B8', 'Total')

# Table Plain D11:D16: no active filter; row 14 hidden by hand.
text('D11', 'Num')
for r, v in zip(range(12, 17), [1, 2, 4, 8, 16]):
    num(f'D{r}', v)

# #This Row from the header row, a data row, the totals row, below and above a table.
for ref, f in [('F1', 'Sales[[#This Row],[Val]]'), ('F2', 'Sales[[#This Row],[Val]]'),
               ('F8', 'Sales[[#This Row],[Val]]'), ('F9', 'Sales[[#This Row],[Val]]'),
               ('F10', 'Plain[[#This Row],[Num]]'), ('F11', 'Plain[[#This Row],[Num]]'),
               ('F12', 'Plain[[#This Row],[Num]]')]:
    formula(ref, f); text('G' + ref[1:], ref + ': ' + f)

cases = ['2*3', 'SUM(Sales[Val])', 'SUBTOTAL(9,Sales[Val])', 'SUBTOTAL(109,Sales[Val])', 'SUBTOTAL(9,A2:A7)']
cases += [f'_xlfn.AGGREGATE(9,{o},Sales[Val])' for o in range(8)]
cases += ['SUM(Plain[Num])', 'SUBTOTAL(9,Plain[Num])', 'SUBTOTAL(109,Plain[Num])']
cases += [f'_xlfn.AGGREGATE(9,{o},Plain[Num])' for o in range(8)]
for i, f in enumerate(cases, start=20):
    text(f'H{i}', f.replace('_xlfn.', '')); formula(f'I{i}', f)

HID = ' hidden="1"'
def colkey(c): return (len(c), c)
sheet_rows = ''.join(
    f'<row r="{r}"{HID if r in hidden else ""}>' + ''.join(rows[r][c] for c in sorted(rows[r], key=colkey)) + '</row>'
    for r in sorted(rows))
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
sheet = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{NS}" xmlns:r="{RNS}">'
         f'<sheetData>{sheet_rows}</sheetData><tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts></worksheet>')
style_info = '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
table1 = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="{NS}" id="1" name="Sales" displayName="Sales" ref="A1:B8" totalsRowCount="1">'
          '<autoFilter ref="A1:B7"><filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn></autoFilter>'
          '<tableColumns count="2"><tableColumn id="1" name="Val" totalsRowFunction="sum"/><tableColumn id="2" name="Grp" totalsRowLabel="Total"/></tableColumns>'
          f'{style_info}</table>')
table2 = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="{NS}" id="2" name="Plain" displayName="Plain" ref="D11:D16" totalsRowShown="0">'
          '<autoFilter ref="D11:D16"/><tableColumns count="1"><tableColumn id="1" name="Num"/></tableColumns>'
          f'{style_info}</table>')
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
            '<sheets><sheet name="Spike" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
def rels(items): return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PR}">'
                         + ''.join(f'<Relationship Id="{i}" Type="{T}{t}" Target="{g}"/>' for i, t, g in items) + '</Relationships>')
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'
content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    f'<Override PartName="/xl/workbook.xml" ContentType="{CT}sheet.main+xml"/>'
    f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{CT}worksheet+xml"/>'
    f'<Override PartName="/xl/tables/table1.xml" ContentType="{CT}table+xml"/>'
    f'<Override PartName="/xl/tables/table2.xml" ContentType="{CT}table+xml"/>'
    f'<Override PartName="/xl/sharedStrings.xml" ContentType="{CT}sharedStrings+xml"/>'
    f'<Override PartName="/xl/styles.xml" ContentType="{CT}styles+xml"/></Types>')

with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', rels([('rId1', 'officeDocument', 'xl/workbook.xml')]))
    z.writestr('xl/workbook.xml', workbook)
    z.writestr('xl/_rels/workbook.xml.rels', rels([('rId1', 'worksheet', 'worksheets/sheet1.xml'),
        ('rId2', 'sharedStrings', 'sharedStrings.xml'), ('rId3', 'styles', 'styles.xml')]))
    z.writestr('xl/worksheets/sheet1.xml', sheet)
    z.writestr('xl/worksheets/_rels/sheet1.xml.rels', rels([('rId1', 'table', '../tables/table1.xml'), ('rId2', 'table', '../tables/table2.xml')]))
    z.writestr('xl/tables/table1.xml', table1)
    z.writestr('xl/tables/table2.xml', table2)
    z.writestr('xl/sharedStrings.xml', sst_xml)
    z.writestr('xl/styles.xml', styles)
print(out)
