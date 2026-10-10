# Saving .xlsx files

`XlsxWriter` saves a workbook as an .xlsx file: one read by `XlsxReader`, with whatever you changed,
or one built in code.

```csharp snippet=save-template
Workbook workbook = XlsxReader.Load(templatePath);
workbook["Budget"]["B2"].Value = 1350;   // an input
workbook.Recalculate();                  // required before saving
XlsxWriter.Save(workbook, resultPath);   // charts, comments and other cells stay as they were
```

```csharp snippet=save-new
var workbook = new Workbook();
Worksheet sheet = workbook.AddSheet("Report");
sheet["A1"].Value = "Total";
sheet["A1"].Style = sheet["A1"].Style with { Font = sheet["A1"].Style.Font with { Bold = true } };
sheet["B1"].Formula = "=SUM(B2:B4)";
sheet["B2"].Value = 10;
sheet["B3"].Value = 20;
sheet["B4"].Value = 30;
sheet["B1"].NumberFormat = "#,##0.00";
sheet.SetColumnWidth(1, 15);
workbook.Recalculate();
XlsxWriter.Save(workbook, path);   // a new .xlsx file; Excel shows Total 60.00
```

## What is written

Everything SharpCell models:

- Cell values and formulas, with the results SharpCell calculated. Formulas are written the way
  Excel stores them, so newer functions such as `XLOOKUP`, `FILTER` or `LET` work when Excel opens
  the file.
- Dynamic arrays, with the area they spill over. A formula that reads a range where one value is
  expected (`=SUM(A1:A3*2)`) is saved as a dynamic array formula, so Excel does not reduce it as
  Excel 2016 would.
- Cell styles and number formats, column widths, hidden columns, row heights, hidden rows, merged
  cells, frozen panes, gridlines, the default column width and row height, and whether a sheet is
  filtered.
- New sheets, defined names, new tables, the date system and the iteration settings.

A workbook read from a file keeps everything SharpCell does not model, copied unchanged from the
file: charts, pivot tables, comments, conditional formats, data validation, pictures, macros,
pattern fills, and parts SharpCell does not know. A cell that did not change is copied as it was,
and a new style starts from the cell's old one, so what SharpCell does not read (rotated text, a
pattern fill) stays. Saving a workbook in which nothing changed writes the original file byte for
byte.

New text goes into the cell itself, not into the shared string table. A new table's header cells
are saved holding the column names as text, as Excel requires. When formulas change, the file's
calculation chain is left out; Excel builds it again. `#SPILL!` and `#CALC!` are saved the way Excel
saves them, so Excel shows them before it recalculates.

After rows or columns were inserted or deleted, everything in the file that points at cells
follows them, including what SharpCell does not model: conditional formats, data validation,
hyperlinks, filters, page breaks, print areas, pictures and shapes, comments, chart series and
pivot table sources. A comment on a deleted cell goes; a picture anchored in deleted rows shrinks to
where they were. A pivot table moves with the rows above it, but rows inserted into it or deleted
from it make the save fail, as Excel refuses that edit. A sheet holding an Excel data table cannot be
saved after its rows or columns moved (`NotSupportedException`). Formulas SharpCell cannot read are
kept as the file has them, with the results Excel saved.

Saving to a path writes a new file next to the target and then replaces it, so the file the workbook
was read from can be the target.

## What cannot be saved

| Situation | Exception |
|---|---|
| A workbook without sheets | `InvalidOperationException` |
| A changed cell of an Excel data table, rows or columns moved on a sheet with one, or `#SPILL!`/`#CALC!` typed as a value | `NotSupportedException` |
| Rows or columns inserted into or deleted from a pivot table | `NotSupportedException` |
| A formula SharpCell cannot read changed so that formulas sharing its text would lose it | `NotSupportedException` |
| Formulas not calculated since the last change (or since loading) | `InvalidOperationException` |
| Formulas SharpCell cannot calculate | `XlsxWriteException` |

Nothing is written when a save is refused.

## Formulas SharpCell cannot calculate

A formula that calls a function SharpCell does not know, or links to another workbook, has no
trustworthy result. By default `Save` refuses with an `XlsxWriteException`; its `Problems` list the
cells. With `new XlsxWriteOptions { KeepUncalculated = true }` those cells keep the result Excel saved,
and the file asks Excel to recalculate every formula when it opens it. Until then, programs that read
saved results (pandas, openpyxl, `XlsxReader`) see what SharpCell calculated for the cells that read
such a formula.

## Memory

A workbook read by `XlsxReader` keeps the file's bytes in memory for saving: count the file's size on
top of the workbook.
