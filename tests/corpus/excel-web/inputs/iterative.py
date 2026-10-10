# Builds the uncalculated inputs of tests/corpus/excel-web/iterative-{1,3,default}.xlsx: the same
# circular references calculated with iterative calculation at 1 and 3 iterations (MaxChange 0,
# so every iteration shows) and at Excel's defaults (100 iterations, MaxChange 0.001).
# Usage: python3 iterative.py <output folder>
import os, sys, zipfile
from xml.sax.saxutils import escape

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

def col_key(ref):
    letters = ''.join(c for c in ref if c.isalpha())
    return (len(letters), letters)

class Sheet:
    def __init__(self, name): self.name, self.rows = name, {}
    def put(self, ref, xml):
        letters = ''.join(c for c in ref if c.isalpha())
        self.rows.setdefault(int(ref[len(letters):]), {})[letters] = xml
    def num(self, ref, v): self.put(ref, f'<c r="{ref}"><v>{v}</v></c>')
    def f(self, ref, formula): self.put(ref, f'<c r="{ref}"><f>{escape(formula)}</f></c>')
    def array(self, ref, area, formula): self.put(ref, f'<c r="{ref}"><f t="array" ref="{area}">{escape(formula)}</f></c>')
    def xml(self):
        rows = ''.join(f'<row r="{r}">' + ''.join(self.rows[r][c] for c in sorted(self.rows[r], key=lambda c: (len(c), c))) + '</row>'
                       for r in sorted(self.rows))
        return f'{HEAD}<worksheet xmlns="{NS}" xmlns:r="{RNS}"><sheetData>{rows}</sheetData></worksheet>'

def sheets():
    out = []
    s = Sheet('Counter')
    s.f('A1', 'A1+1')                                 # self-counter
    s.f('B1', 'B1+1'); s.f('B2', 'B1*2')              # a reader outside the cycle
    s.f('C1', 'SUM(C1:C2)+1'); s.num('C2', 5)         # self-reference through a range
    out.append(s)
    s = Sheet('Pairs')
    s.f('A1', 'B1+1'); s.f('B1', 'A1*2')              # across a row, first cell reads the second
    s.f('A3', 'A4*2'); s.f('A4', 'A3+1')              # down a column
    s.f('D1', 'E1*2'); s.f('E1', 'D1+1')              # the same pair, the formulas swapped
    s.f('G1', 'H2+1'); s.f('H2', 'G1*2'); s.f('G3', 'G1+H2')   # diagonal, with a reader
    out.append(s)
    s = Sheet('Sheets'); s.f('A1', 'Other!A1+1'); out.append(s)
    s = Sheet('Model')
    s.num('A1', 1000); s.f('A2', 'A1+A3'); s.f('A3', '(A1+A2)/2*0.05'); s.f('B1', 'A2')
    out.append(s)
    s = Sheet('Diverging'); s.f('A1', 'A1*2+1'); out.append(s)
    s = Sheet('Array'); s.f('A1', 'A2+1'); s.array('A2', 'A2:A3', 'A1'); out.append(s)
    s = Sheet('Other'); s.f('A1', 'Sheets!A1*2'); out.append(s)
    return out

def write(path, count, delta):
    sh = sheets()
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        def rels(items): return f'{HEAD}<Relationships xmlns="{PR}">' + ''.join(f'<Relationship Id="{i}" Type="{T}{t}" Target="{g}"/>' for i, t, g in items) + '</Relationships>'
        overrides = [('/xl/workbook.xml', 'sheet.main+xml')]
        for n, s in enumerate(sh, start=1):
            z.writestr(f'xl/worksheets/sheet{n}.xml', s.xml())
            overrides.append((f'/xl/worksheets/sheet{n}.xml', 'worksheet+xml'))
        sheet_list = ''.join(f'<sheet name="{s.name}" sheetId="{n}" r:id="rId{n}"/>' for n, s in enumerate(sh, start=1))
        calc = f'<calcPr calcId="191029" fullCalcOnLoad="1" iterate="1" iterateCount="{count}" iterateDelta="{delta}"/>'
        z.writestr('xl/workbook.xml', f'{HEAD}<workbook xmlns="{NS}" xmlns:r="{RNS}"><sheets>{sheet_list}</sheets>{calc}</workbook>')
        z.writestr('xl/_rels/workbook.xml.rels', rels([(f'rId{n}', 'worksheet', f'worksheets/sheet{n}.xml') for n in range(1, len(sh) + 1)]))
        z.writestr('_rels/.rels', rels([('rId1', 'officeDocument', 'xl/workbook.xml')]))
        z.writestr('[Content_Types].xml', f'{HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                   + ''.join(f'<Override PartName="{p}" ContentType="{CT}{c}"/>' for p, c in overrides) + '</Types>')
    print(path)

folder = sys.argv[1]
write(os.path.join(folder, 'iterative-1.xlsx'), 1, 0)
write(os.path.join(folder, 'iterative-3.xlsx'), 3, 0)
write(os.path.join(folder, 'iterative-default.xlsx'), 100, 0.001)
