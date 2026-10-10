# Saving .xlsx files

`XlsxWriter` saves a workbook back into the file it was read from, with the values SharpCell
calculated. It is made for templates: read a workbook, change input cells, recalculate, save.

```csharp snippet=save-template
Workbook workbook = XlsxReader.Load(templatePath);
workbook["Budget"]["B2"].Value = 1350;   // an input
workbook.Recalculate();                  // required before saving
XlsxWriter.Save(workbook, resultPath);   // styles, charts and other cells stay as they were
```

## What is written

Only cell values. Everything else in the file is copied unchanged: styles, number formats, column
widths, charts, pivot tables, comments, data validation, macros and parts SharpCell does not know.
A cell whose value did not change is copied as it was. New text goes into the cell itself, not into
the shared string table.

Dynamic arrays that grow or shrink get their new cells and area. `#SPILL!` and `#CALC!` are saved the
way Excel saves them, so Excel shows them before it recalculates.

Saving a workbook with no changed values writes the original file byte for byte. Saving to a path
writes a new file next to the target and then replaces it, so the file the workbook was read from can
be the target.

## What cannot be saved

| Situation | Exception |
|---|---|
| A workbook built in code, not read by `XlsxReader` | `NotSupportedException` |
| A formula added, removed or changed, or a value typed over a formula | `NotSupportedException` |
| Sheets, defined names, tables, hidden rows, filters or the date system changed | `NotSupportedException` |
| A number format changed or set | `NotSupportedException` |
| A changed cell of an Excel data table, or `#SPILL!`/`#CALC!` typed as a value | `NotSupportedException` |
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
