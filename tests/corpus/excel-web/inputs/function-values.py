# Builds the uncalculated input of tests/corpus/excel-web/function-values.xlsx: built-in function
# names passed as values ("eta-reduced lambdas", =BYROW(A1:B3,SUM)), which files store with the
# _xleta. prefix. Each formula sits in column F, its label in E, four rows apart so spills fit.
import sys, zipfile
from xml.sax.saxutils import escape

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
METADATA = (f'{HEAD}<metadata xmlns="{NS}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
            '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" '
            'splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
            '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
            '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
            '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>')

# (label as typed in Excel, formula as a file stores it)
CASES = [
    ('BYROW(A1:B3,SUM)', '_xlfn.BYROW(A1:B3,_xleta.SUM)'),
    ('BYCOL(A1:B3,MAX)', '_xlfn.BYCOL(A1:B3,_xleta.MAX)'),
    ('MAP(D1:D3,LEN)', '_xlfn.MAP(D1:D3,_xleta.LEN)'),
    ('MAP(A1:A3,B1:B3,SUM)', '_xlfn.MAP(A1:A3,B1:B3,_xleta.SUM)'),
    ('REDUCE(0,A1:B3,SUM)', '_xlfn.REDUCE(0,A1:B3,_xleta.SUM)'),
    ('SCAN(0,A1:A3,SUM)', '_xlfn.SCAN(0,A1:A3,_xleta.SUM)'),
    ('MAKEARRAY(2,2,SUM)', '_xlfn.MAKEARRAY(2,2,_xleta.SUM)'),
    ('LET(f,SUM,f(1,2,3))', '_xlfn.LET(_xlpm.f,_xleta.SUM,_xlpm.f(1,2,3))'),
    ('LET(f,LEN,f("abcd"))', '_xlfn.LET(_xlpm.f,_xleta.LEN,_xlpm.f("abcd"))'),
    ('MAP(A1:A3,ABS)', '_xlfn.MAP(A1:A3,_xleta.ABS)'),
    ('MAP({-1,"x",TRUE},ABS)', '_xlfn.MAP({-1,"x",TRUE},_xleta.ABS)'),
    ('BYROW(A1:B3,AVERAGE)', '_xlfn.BYROW(A1:B3,_xleta.AVERAGE)'),
    ('BYROW(A1:B3,COUNTA)', '_xlfn.BYROW(A1:B3,_xleta.COUNTA)'),
    ('MAP(D1:D3,UPPER)', '_xlfn.MAP(D1:D3,_xleta.UPPER)'),
    ('REDUCE("",D1:D3,CONCAT)', '_xlfn.REDUCE("",D1:D3,_xleta.CONCAT)'),
    ('BYROW(A1:B3,IF)', '_xlfn.BYROW(A1:B3,_xleta.IF)'),
    ('MAP(A1:A3,ISOMITTED)', '_xlfn.MAP(A1:A3,_xleta.ISOMITTED)'),
    ('LET(f,ABS,f(1,2))', '_xlfn.LET(_xlpm.f,_xleta.ABS,_xlpm.f(1,2))'),
    ('LET(f,ABS,f())', '_xlfn.LET(_xlpm.f,_xleta.ABS,_xlpm.f())'),
    ('LET(f,SUM,f)', '_xlfn.LET(_xlpm.f,_xleta.SUM,_xlpm.f)'),
    ('=SUM written plainly', 'SUM'),
    ('=SUM written as eta', '_xleta.SUM'),
    ('BYROW(A1:B3,NOSUCHFN)', '_xlfn.BYROW(A1:B3,NOSUCHFN)'),
    ('BYROW(A1:B3,Rate) with Rate a name', '_xlfn.BYROW(A1:B3,Rate)'),
    ('BYROW(A1:B3,LAMBDA(r,SUM(r)))', '_xlfn.BYROW(A1:B3,_xlfn.LAMBDA(_xlpm.r,SUM(_xlpm.r)))'),
    ('ISOMITTED(SUM)', '_xlfn.ISOMITTED(_xleta.SUM)'),
    ('TYPE(SUM)', 'TYPE(_xleta.SUM)'),
    ('ISERROR(SUM)', 'ISERROR(_xleta.SUM)'),
]

strings = []
def sst(s):
    if s not in strings: strings.append(s)
    return strings.index(s)

rows = {}
def put(column, r, xml): rows.setdefault(r, {})[column] = xml
for r, (a, b) in enumerate([(1, 2), (3, 4), (5, 6)], start=1):
    put('A', r, f'<c r="A{r}"><v>{a}</v></c>')
    put('B', r, f'<c r="B{r}"><v>{b}</v></c>')
for r, t in enumerate(['ab', 'cde', 'f'], start=1):
    put('D', r, f'<c r="D{r}" t="s"><v>{sst(t)}</v></c>')
for i, (label, formula) in enumerate(CASES):
    r = 1 + 4 * i
    put('E', r, f'<c r="E{r}" t="s"><v>{sst(label)}</v></c>')
    # A dynamic array formula, as Excel marks it: cm points at the XLDAPR cell metadata. Without
    # it Excel reads the formula as one written before dynamic arrays and does not spill it.
    put('F', r, f'<c r="F{r}" cm="1"><f t="array" ref="F{r}">{escape(formula)}</f></c>')
sheet_rows = ''.join(f'<row r="{r}">' + ''.join(rows[r][c] for c in sorted(rows[r])) + '</row>' for r in sorted(rows))
sheet = f'{HEAD}<worksheet xmlns="{NS}" xmlns:r="{RNS}"><sheetData>{sheet_rows}</sheetData></worksheet>'

out = sys.argv[1]
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    def rels(items): return f'{HEAD}<Relationships xmlns="{PR}">' + ''.join(f'<Relationship Id="{i}" Type="{T}{t}" Target="{g}"/>' for i, t, g in items) + '</Relationships>'
    z.writestr('xl/worksheets/sheet1.xml', sheet)
    z.writestr('xl/workbook.xml', f'{HEAD}<workbook xmlns="{NS}" xmlns:r="{RNS}"><sheets><sheet name="F" sheetId="1" r:id="rId1"/></sheets>'
               '<definedNames><definedName name="Rate">0.5</definedName></definedNames><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    z.writestr('xl/_rels/workbook.xml.rels', rels([('rId1', 'worksheet', 'worksheets/sheet1.xml'), ('rId2', 'sharedStrings', 'sharedStrings.xml'), ('rId3', 'sheetMetadata', 'metadata.xml')]))
    z.writestr('xl/metadata.xml', METADATA)
    z.writestr('_rels/.rels', rels([('rId1', 'officeDocument', 'xl/workbook.xml')]))
    z.writestr('xl/sharedStrings.xml', f'{HEAD}<sst xmlns="{NS}" count="{len(strings)}" uniqueCount="{len(strings)}">' + ''.join(f'<si><t xml:space="preserve">{escape(s)}</t></si>' for s in strings) + '</sst>')
    overrides = [('/xl/workbook.xml', 'sheet.main+xml'), ('/xl/worksheets/sheet1.xml', 'worksheet+xml'), ('/xl/sharedStrings.xml', 'sharedStrings+xml'), ('/xl/metadata.xml', 'sheetMetadata+xml')]
    z.writestr('[Content_Types].xml', f'{HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
               + ''.join(f'<Override PartName="{p}" ContentType="{CT}{c}"/>' for p, c in overrides) + '</Types>')
print(out)
