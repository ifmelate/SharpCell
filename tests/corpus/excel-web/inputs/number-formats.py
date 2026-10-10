# Builds the uncalculated input of tests/corpus/excel-web/number-formats.xlsx: TEXT(value, code)
# next to the same value shown through a cell style with that code, for Cell.Text. Excel saves
# TEXT's result but not how a cell is shown, so Cell.Text of column C is checked against column B.
import sys, zipfile
from xml.sax.saxutils import escape

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

# (value, code): a value is a number, text, a bool, '=formula' or None for an empty cell.
CASES = [
    (1234.5678, '0'), (1234.5678, '0.00'), (1234.5678, '#,##0.00'), (1234.5678, '0%'),
    (1234.5678, '0.00E+00'), (1234.5678, '# ?/?'), (1234.5678, 'General'),
    (-1234.5, '#,##0.00;(#,##0.00)'), (-1234.5, '#,##0 ;[Red](#,##0)'), (-1234.5, '0.00;-0.00;"zero"'),
    (0, '0.00;-0.00;"zero"'), (0, 'General'),
    (46096.5, 'yyyy-mm-dd'), (46096.5, 'd-mmm-yy'), (46096.5, 'h:mm AM/PM'), (46096.5, 'm/d/yy h:mm'),
    (46096.5, '[h]:mm:ss'), (46096.5, 'mm-dd-yy'), (46096.5, 'dddd, mmmm d, yyyy'),
    (-1, 'yyyy-mm-dd'),
    (0.000012345, 'General'), (1.5e21, 'General'), (123456789012, 'General'), (0.1 + 0.2, 'General'),
    ('hello', '@'), ('hello', '"Name: "@'), ('hello', '0.00'), ('12', '0.00'),
    (True, '0.00'), ('=1/0', '0.00'), (None, '0.00'),
]

strings = []
def sst(s):
    if s not in strings: strings.append(s)
    return strings.index(s)

def number(v): return repr(float(v)) if isinstance(v, float) else str(v)

codes = []   # custom codes, numFmtId 164 + index; cellXfs 0 is General, 1 + index uses code index
def style_of(code):
    if code == 'General': return 0
    if code not in codes: codes.append(code)
    return 1 + codes.index(code)

def value_cell(ref, v, style=None):
    s = f' s="{style}"' if style else ''
    if v is None: return f'<c r="{ref}"{s}/>' if style else ''
    if isinstance(v, bool): return f'<c r="{ref}"{s} t="b"><v>{int(v)}</v></c>'
    if isinstance(v, str) and v.startswith('='): return f'<c r="{ref}"{s}><f>{escape(v[1:])}</f></c>'
    if isinstance(v, str): return f'<c r="{ref}"{s} t="s"><v>{sst(v)}</v></c>'
    return f'<c r="{ref}"{s}><v>{number(v)}</v></c>'

rows = []
for r, (v, code) in enumerate(CASES, start=1):
    cells = value_cell(f'A{r}', v)
    cells += f'<c r="B{r}"><f>TEXT(A{r},D{r})</f></c>'
    cells += value_cell(f'C{r}', v, style_of(code))
    cells += f'<c r="D{r}" t="s"><v>{sst(code)}</v></c>'
    rows.append(f'<row r="{r}">{cells}</row>')
sheet = f'{HEAD}<worksheet xmlns="{NS}" xmlns:r="{RNS}"><sheetData>{"".join(rows)}</sheetData></worksheet>'

num_fmts = ''.join(f'<numFmt numFmtId="{164 + i}" formatCode="{escape(c, {chr(34): "&quot;"})}"/>' for i, c in enumerate(codes))
xfs = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' + ''.join(
    f'<xf numFmtId="{164 + i}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' for i in range(len(codes)))
styles = (f'{HEAD}<styleSheet xmlns="{NS}"><numFmts count="{len(codes)}">{num_fmts}</numFmts>'
          '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          f'<cellXfs count="{1 + len(codes)}">{xfs}</cellXfs>'
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

out = sys.argv[1]
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    def rels(items): return f'{HEAD}<Relationships xmlns="{PR}">' + ''.join(f'<Relationship Id="{i}" Type="{T}{t}" Target="{g}"/>' for i, t, g in items) + '</Relationships>'
    z.writestr('xl/worksheets/sheet1.xml', sheet)
    z.writestr('xl/workbook.xml', f'{HEAD}<workbook xmlns="{NS}" xmlns:r="{RNS}"><sheets><sheet name="Formats" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    z.writestr('xl/_rels/workbook.xml.rels', rels([('rId1', 'worksheet', 'worksheets/sheet1.xml'), ('rId2', 'sharedStrings', 'sharedStrings.xml'), ('rId3', 'styles', 'styles.xml')]))
    z.writestr('_rels/.rels', rels([('rId1', 'officeDocument', 'xl/workbook.xml')]))
    z.writestr('xl/sharedStrings.xml', f'{HEAD}<sst xmlns="{NS}" count="{len(strings)}" uniqueCount="{len(strings)}">' + ''.join(f'<si><t xml:space="preserve">{escape(s)}</t></si>' for s in strings) + '</sst>')
    z.writestr('xl/styles.xml', styles)
    overrides = [('/xl/workbook.xml', 'sheet.main+xml'), ('/xl/worksheets/sheet1.xml', 'worksheet+xml'), ('/xl/sharedStrings.xml', 'sharedStrings+xml'), ('/xl/styles.xml', 'styles+xml')]
    z.writestr('[Content_Types].xml', f'{HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
               + ''.join(f'<Override PartName="{p}" ContentType="{CT}{c}"/>' for p, c in overrides) + '</Types>')
print(out)
