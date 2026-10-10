# Builds the uncalculated input of tests/corpus/excel-web/groupby.xlsx: GROUPBY, PIVOTBY and
# PERCENTOF over one small table, every option Microsoft documents, and the errors. Each case is a
# label in column A and the formula in column B, 25 rows apart so the results can spill.
import re, sys, zipfile
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

# Region, Product, Units, Price; one blank product and one #N/A in Units.
DATA = [
    ('North', 'A', 10, 1.5), ('South', 'B', 20, 2.25), ('North', 'A', 30, 1.5), ('South', 'C', 40, 3),
    ('North', 'B', 50, 2.25), ('South', 'A', 60, 1.5), ('North', 'C', 70, 3), ('South', 'B', 80, 2.25),
    ('North', 'A', 90, 1.5), ('South', 'C', 100, 3), ('North', None, 110, 1.5), ('South', 'B', '#N/A', 2.25),
]

R, P, U, PR_, U2 = 'Data!A2:A13', 'Data!B2:B13', 'Data!C2:C13', 'Data!D2:D13', 'Data!C2:D13'
RH, PH, UH = 'Data!A1:A13', 'Data!B1:B13', 'Data!C1:C13'
RP = 'Data!A2:B13'
GROUPBY = [
    f'GROUPBY({R},{U},SUM)',
    *[f'GROUPBY({RH},{UH},SUM,{h})' for h in range(4)],
    f'GROUPBY({RH},{UH},SUM)',
    *[f'GROUPBY({RP},{U},SUM,0,{d})' for d in (-2, -1, 0, 1, 2)],
    f'GROUPBY({RP},{U},SUM)',
    f'GROUPBY({RH},{UH},SUM,3,2)',
    *[f'GROUPBY({R},{U},SUM,0,1,{s})' for s in (-1, 2, -2)],
    f'GROUPBY({RP},{U},SUM,0,1,{{1,-2}})',
    f'GROUPBY({RP},{U},SUM,0,1,{{-1,3}})',
    f'GROUPBY({R},{U},SUM,0,1,1,{R}="North")',
    f'GROUPBY({RP},{U},SUM,0,1,,,1)',
    f'GROUPBY({RP},{U},SUM,0,2,,,1)',
    *[f'GROUPBY({R},{U},{fn})' for fn in ('AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'PERCENTOF', 'HSTACK(SUM,AVERAGE)', 'LAMBDA(x,MAX(x)-MIN(x))')],
    f'GROUPBY({R},{U2},SUM)',
    f'GROUPBY({P},{U},SUM)',
    f'GROUPBY({P},{PR_},SUM)',
    f'GROUPBY({R},{U},SUM,0,1,1,Data!A2:A12="North")',
    f'GROUPBY({R},Data!C2:C12,SUM)',
    f'GROUPBY({R},{U},1)',
    f'GROUPBY({R},{U},SUM,7)',
    f'GROUPBY({R},{U},SUM,0,3)',
    f'GROUPBY({R},{U},SUM,0,1,5)',
    f'GROUPBY({R},{U},SUM,0,1,1,{R}="East")',
]
PIVOTBY = [
    f'PIVOTBY({R},{P},{U},SUM)',
    *[f'PIVOTBY({RH},{PH},{UH},SUM,{h})' for h in range(4)],
    *[f'PIVOTBY({R},{P},{PR_},SUM,0,{d})' for d in (0, -1, 2)],
    *[f'PIVOTBY({R},{P},{PR_},SUM,0,1,,{d})' for d in (0, -1)],
    f'PIVOTBY({R},{P},{PR_},SUM,0,1,-1)',
    f'PIVOTBY({R},{P},{PR_},SUM,0,1,1,1,-1)',
    *[f'PIVOTBY({R},{P},{PR_},PERCENTOF,0,1,1,1,1,,{k})' for k in range(5)],
    f'PIVOTBY({R},{P},{PR_},SUM,0,1,1,1,1,{R}="North")',
    f'PIVOTBY({R},{P},Data!C2:D13,SUM)',
    f'PIVOTBY({RP},Data!A2:A13,{PR_},SUM)',
    f'PIVOTBY({R},{P},{PR_},AVERAGE)',
    f'PIVOTBY({R},Data!B2:B12,{PR_},SUM)',
    f'PIVOTBY({R},{P},{PR_},SUM,0,1,1,1,1,,9)',
]
PERCENT = [
    'PERCENTOF(Data!D2:D7,Data!D2:D13)',
    f'PERCENTOF({U},{U})',
    'PERCENTOF(Data!B2:B4,Data!D2:D4)',
    'PERCENTOF({1,2},{0,0})',
    'PERCENTOF(5,10)',
    'PERCENTOF(Data!D2:D7,Data!C2:C11)',
]

FUTURE = ('GROUPBY', 'PIVOTBY', 'PERCENTOF', 'HSTACK', 'LAMBDA')
ETA = ('SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'PERCENTOF')

def to_file(formula):
    """The formula as a file stores it: _xlfn. on newer functions, _xleta. on functions passed as values."""
    out = re.sub(r'\b(' + '|'.join(FUTURE) + r')\(', r'_xlfn.\1(', formula)
    out = re.sub(r'(?<![\w.])(' + '|'.join(ETA) + r')(?=[,)])', r'_xleta.\1', out)
    return out.replace('_xlfn.LAMBDA(x,MAX(x)-MIN(x))', '_xlfn.LAMBDA(_xlpm.x,MAX(_xlpm.x)-MIN(_xlpm.x))')

strings = []
def sst(s):
    if s not in strings: strings.append(s)
    return strings.index(s)

def column_key(ref):
    letters = ''.join(c for c in ref if c.isalpha())
    return (len(letters), letters)

class Sheet:
    def __init__(self, name): self.name, self.rows = name, {}
    def put(self, ref, xml):
        letters = ''.join(c for c in ref if c.isalpha())
        self.rows.setdefault(int(ref[len(letters):]), {})[letters] = xml
    def value(self, ref, v):
        if v is None: return
        if isinstance(v, str) and v.startswith('#'): self.put(ref, f'<c r="{ref}" t="e"><v>{v}</v></c>')
        elif isinstance(v, str): self.put(ref, f'<c r="{ref}" t="s"><v>{sst(v)}</v></c>')
        else: self.put(ref, f'<c r="{ref}"><v>{v}</v></c>')
    def cases(self, formulas):
        for i, f in enumerate(formulas):
            r = 1 + 25 * i
            self.value(f'A{r}', f)
            self.put(f'B{r}', f'<c r="B{r}" cm="1"><f t="array" ref="B{r}">{escape(to_file(f))}</f></c>')
    def xml(self):
        rows = ''.join(f'<row r="{r}">' + ''.join(self.rows[r][c] for c in sorted(self.rows[r], key=lambda c: (len(c), c))) + '</row>'
                       for r in sorted(self.rows))
        return f'{HEAD}<worksheet xmlns="{NS}" xmlns:r="{RNS}"><sheetData>{rows}</sheetData></worksheet>'

data = Sheet('Data')
for c, h in zip('ABCD', ('Region', 'Product', 'Units', 'Price')): data.value(f'{c}1', h)
for r, row in enumerate(DATA, start=2):
    for c, v in zip('ABCD', row): data.value(f'{c}{r}', v)
g, p, pct = Sheet('G'), Sheet('P'), Sheet('Pct')
g.cases(GROUPBY); p.cases(PIVOTBY); pct.cases(PERCENT)
sheets = [data, g, p, pct]

out = sys.argv[1]
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    def rels(items): return f'{HEAD}<Relationships xmlns="{PR}">' + ''.join(f'<Relationship Id="{i}" Type="{T}{t}" Target="{g}"/>' for i, t, g in items) + '</Relationships>'
    overrides = [('/xl/workbook.xml', 'sheet.main+xml'), ('/xl/sharedStrings.xml', 'sharedStrings+xml'), ('/xl/metadata.xml', 'sheetMetadata+xml')]
    for n, s in enumerate(sheets, start=1):
        z.writestr(f'xl/worksheets/sheet{n}.xml', s.xml())
        overrides.append((f'/xl/worksheets/sheet{n}.xml', 'worksheet+xml'))
    sheet_list = ''.join(f'<sheet name="{s.name}" sheetId="{n}" r:id="rId{n}"/>' for n, s in enumerate(sheets, start=1))
    z.writestr('xl/workbook.xml', f'{HEAD}<workbook xmlns="{NS}" xmlns:r="{RNS}"><sheets>{sheet_list}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    k = len(sheets)
    z.writestr('xl/_rels/workbook.xml.rels', rels([(f'rId{n}', 'worksheet', f'worksheets/sheet{n}.xml') for n in range(1, k + 1)]
                                                  + [(f'rId{k + 1}', 'sharedStrings', 'sharedStrings.xml'), (f'rId{k + 2}', 'sheetMetadata', 'metadata.xml')]))
    z.writestr('_rels/.rels', rels([('rId1', 'officeDocument', 'xl/workbook.xml')]))
    z.writestr('xl/sharedStrings.xml', f'{HEAD}<sst xmlns="{NS}" count="{len(strings)}" uniqueCount="{len(strings)}">' + ''.join(f'<si><t xml:space="preserve">{escape(s)}</t></si>' for s in strings) + '</sst>')
    z.writestr('xl/metadata.xml', METADATA)
    z.writestr('[Content_Types].xml', f'{HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
               + ''.join(f'<Override PartName="{p}" ContentType="{CT}{c}"/>' for p, c in overrides) + '</Types>')
print(out)
