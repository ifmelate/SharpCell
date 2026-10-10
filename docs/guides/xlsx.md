# Reading .xlsx files

What `XlsxReader` loads, what it skips, and how saved results and streams behave.

## What is read

- Sheets, including hidden ones, in workbook order.
- Constants: numbers, text, logical values and errors.
- Formulas: ordinary, shared, legacy array formulas (Ctrl+Shift+Enter) and dynamic arrays.
- The results Excel saved with each formula.
- Defined names, workbook-wide and sheet-scoped.
- Tables, with their column names and header and totals rows.
- Hidden rows, and whether a sheet has a filter (see [Tables and hidden rows](tables.md)).
- The date system, 1900 or 1904.

## What is not read

Styles and number formats, charts, pivot tables, comments and macros are skipped.

| File | Result |
|---|---|
| .xls, .xlsm, .xlsb, encrypted files | `NotSupportedException` |
| A damaged or non-xlsx file | `InvalidDataException` |

## Saved results

After `Load`, a cell with a formula holds the result Excel saved in the file. A file written by another program may have no saved results; those cells are empty. `Workbook.Recalculate` calculates every formula with SharpCell and replaces the saved results.

A formula SharpCell cannot parse, such as a link to another workbook, does not fail the load. Once calculated, it is `#NAME?` and listed in `Workbook.Diagnostics` as `UnsupportedFormula`.

A legacy array formula covers a fixed range. Setting the value or formula of one of its cells other than the top-left one throws `InvalidOperationException`; change the whole array at its top-left cell.

## Streams

```csharp snippet=load-stream
using Stream stream = File.OpenRead(path);
Workbook workbook = XlsxReader.Load(stream);   // the stream is left open
```

The stream does not have to be seekable; a stream that is not is copied into memory first.

## Writing

`XlsxWriter` saves new values into the file a workbook was read from; see [Saving .xlsx files](saving.md).
