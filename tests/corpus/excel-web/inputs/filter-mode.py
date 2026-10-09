# Builds the uncalculated input of tests/corpus/excel-web/filter-mode.xlsx: which hidden rows Excel treats as filtered when it opens a file.
import sys, zipfile
from xml.sax.saxutils import escape

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
STYLE = '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
strings = []
def sst(s):
    if s not in strings: strings.append(s)
    return strings.index(s)

class Sheet:
    def __init__(self, name):
        self.name, self.rows, self.hidden, self.tables, self.auto_filter = name, {}, set(), [], ''
    def put(self, ref, xml):
        col = ''.join(c for c in ref if c.isalpha()); self.rows.setdefault(int(ref[len(col):]), {})[col] = xml
    def num(self, ref, v): self.put(ref, f'<c r="{ref}"><v>{v}</v></c>')
    def text(self, ref, s): self.put(ref, f'<c r="{ref}" t="s"><v>{sst(s)}</v></c>')
    def formula(self, ref, f): self.put(ref, f'<c r="{ref}"><f>{escape(f)}</f></c>')
    def block(self, top, label, out_row):
        # 1, 2, 4, 8, 16 in A{top}..A{top+4}, the 4 hidden by hand: sum 31, without it 27.
        for r, v in zip(range(top, top + 5), [1, 2, 4, 8, 16]): self.num(f'A{r}', v)
        self.hidden.add(top + 2)
        rng = f'A{top}:A{top + 4}'
        for i, f in enumerate([f'SUBTOTAL(9,{rng})', f'SUBTOTAL(109,{rng})']):
            self.text(f'C{out_row + i}', f'{label}: {f}'); self.formula(f'D{out_row + i}', f)
    def filtered_table(self, name, criteria_values):
        # Sales A1:B7: Grp x/y/x/y/z/x; the filter keeps criteria_values and hides the other rows.
        self.text('A1', 'Val'); self.text('B1', 'Grp')
        for r, (v, g) in enumerate([(10, 'x'), (20, 'y'), (30, 'x'), (40, 'y'), (50, 'z'), (60, 'x')], start=2):
            self.num(f'A{r}', v); self.text(f'B{r}', g)
            if g not in criteria_values: self.hidden.add(r)
        filters = ''.join(f'<filter val="{v}"/>' for v in criteria_values)
        self.tables.append((name, 'A1:B7', f'<autoFilter ref="A1:B7"><filterColumn colId="1"><filters>{filters}</filters></filterColumn></autoFilter>'
                            '<tableColumns count="2"><tableColumn id="1" name="Val"/><tableColumn id="2" name="Grp"/></tableColumns>'))
        self.text('C9', f'{name}: SUBTOTAL(9,A2:A7)'); self.formula('D9', 'SUBTOTAL(9,A2:A7)')
    def plain_table(self, name, top):
        self.text(f'A{top}', 'Num')
        self.tables.append((name, f'A{top}:A{top + 5}', f'<autoFilter ref="A{top}:A{top + 5}"/><tableColumns count="1"><tableColumn id="1" name="Num"/></tableColumns>'))
    def xml(self, first_table_id):
        HID = ' hidden="1"'
        rows = ''.join(f'<row r="{r}"{HID if r in self.hidden else ""}>' + ''.join(self.rows[r][c] for c in sorted(self.rows[r], key=lambda c: (len(c), c))) + '</row>'
                       for r in sorted(self.rows))
        parts = ''.join(f'<tablePart r:id="rId{i + 1}"/>' for i in range(len(self.tables)))
        parts = f'<tableParts count="{len(self.tables)}">{parts}</tableParts>' if self.tables else ''
        return f'{HEAD}<worksheet xmlns="{NS}" xmlns:r="{RNS}"><sheetData>{rows}</sheetData>{self.auto_filter}{parts}</worksheet>'

sheets = []
# S1: active table filter; hidden rows outside any autoFilter and inside an inactive one.
s = Sheet('ActiveTableFilter'); s.filtered_table('Sales', ['x'])
s.block(20, 'outside any autoFilter', 40); s.plain_table('Plain', 29); s.block(30, 'inside inactive table autoFilter', 43)
sheets.append(s)
# S2: no filter anywhere (control).
s = Sheet('NoFilter'); s.block(20, 'no filter on sheet', 40); sheets.append(s)
# S3: active sheet-level autoFilter on A1:B7; hidden row outside it.
s = Sheet('ActiveSheetFilter')
s.text('A1', 'Val'); s.text('B1', 'Grp')
for r, (v, g) in enumerate([(10, 'x'), (20, 'y'), (30, 'x'), (40, 'y'), (50, 'z'), (60, 'x')], start=2):
    s.num(f'A{r}', v); s.text(f'B{r}', g)
    if g != 'x': s.hidden.add(r)
s.auto_filter = '<autoFilter ref="A1:B7"><filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn></autoFilter>'
s.text('C9', 'sheet filter: SUBTOTAL(9,A2:A7)'); s.formula('D9', 'SUBTOTAL(9,A2:A7)')
s.block(20, 'outside active sheet autoFilter', 40); sheets.append(s)
# S4: table filter with criteria that keep every row; hidden row elsewhere.
s = Sheet('FilterHidesNothing'); s.filtered_table('Everything', ['x', 'y', 'z'])
s.block(20, 'outside filter that hides nothing', 40); sheets.append(s)

out = sys.argv[1]
table_id = 0
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    def rels(items): return f'{HEAD}<Relationships xmlns="{PR}">' + ''.join(f'<Relationship Id="{i}" Type="{T}{t}" Target="{g}"/>' for i, t, g in items) + '</Relationships>'
    overrides = [('/xl/workbook.xml', 'sheet.main+xml'), ('/xl/sharedStrings.xml', 'sharedStrings+xml'), ('/xl/styles.xml', 'styles+xml')]
    defined = []
    for n, sh in enumerate(sheets, start=1):
        table_rels = []
        for i, (name, ref, inner) in enumerate(sh.tables, start=1):
            table_id += 1
            z.writestr(f'xl/tables/table{table_id}.xml', f'{HEAD}<table xmlns="{NS}" id="{table_id}" name="{name}" displayName="{name}" ref="{ref}" totalsRowShown="0">{inner}{STYLE}</table>')
            overrides.append((f'/xl/tables/table{table_id}.xml', 'table+xml'))
            table_rels.append((f'rId{i}', 'table', f'../tables/table{table_id}.xml'))
        z.writestr(f'xl/worksheets/sheet{n}.xml', sh.xml(table_id))
        if table_rels: z.writestr(f'xl/worksheets/_rels/sheet{n}.xml.rels', rels(table_rels))
        overrides.append((f'/xl/worksheets/sheet{n}.xml', 'worksheet+xml'))
        if sh.auto_filter: defined.append(f'<definedName name="_xlnm._FilterDatabase" localSheetId="{n - 1}" hidden="1">{sh.name}!$A$1:$B$7</definedName>')
    sheet_list = ''.join(f'<sheet name="{sh.name}" sheetId="{n}" r:id="rId{n}"/>' for n, sh in enumerate(sheets, start=1))
    names = f'<definedNames>{"".join(defined)}</definedNames>' if defined else ''
    z.writestr('xl/workbook.xml', f'{HEAD}<workbook xmlns="{NS}" xmlns:r="{RNS}"><sheets>{sheet_list}</sheets>{names}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    k = len(sheets)
    z.writestr('xl/_rels/workbook.xml.rels', rels([(f'rId{n}', 'worksheet', f'worksheets/sheet{n}.xml') for n in range(1, k + 1)]
                                                  + [(f'rId{k + 1}', 'sharedStrings', 'sharedStrings.xml'), (f'rId{k + 2}', 'styles', 'styles.xml')]))
    z.writestr('_rels/.rels', rels([('rId1', 'officeDocument', 'xl/workbook.xml')]))
    z.writestr('xl/sharedStrings.xml', f'{HEAD}<sst xmlns="{NS}" count="{len(strings)}" uniqueCount="{len(strings)}">' + ''.join(f'<si><t xml:space="preserve">{escape(s)}</t></si>' for s in strings) + '</sst>')
    z.writestr('xl/styles.xml', f'{HEAD}<styleSheet xmlns="{NS}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
               '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
               '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
               '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
               '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
               '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    z.writestr('[Content_Types].xml', f'{HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
               + ''.join(f'<Override PartName="{p}" ContentType="{CT}{c}"/>' for p, c in overrides) + '</Types>')
print(out)
