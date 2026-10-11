# Roadmap

Where SharpCell is going, in the order the work is planned. Each step ships when its tests pass
and its results are checked against Microsoft Excel. Dates are not promised; the project has one
maintainer.

## Released

| Version | What |
|---|---|
| 0.1 | The engine: parser, dependency graph, recalculation, dynamic arrays, LET/LAMBDA, reading .xlsx |
| 0.2 | Tables and structured references (`Sales[Units]`), hidden rows in `SUBTOTAL` and `AGGREGATE` |
| 0.3 | Saving recalculated values back into the file a workbook was read from |
| 0.4 | API for applications: walking and writing cells, custom functions, changed cells, cloning templates |
| 0.5 | Number formats and `Cell.Text`; iterative calculation of circular references |
| 0.6 | Cell styles, column widths, row heights, merged cells, frozen panes and gridlines, read from files |
| 0.7 | Saving everything the model holds, workbooks built in code, .xlsm files, inserting and deleting rows and columns; functions as values, `GROUPBY`, `PIVOTBY`, `PERCENTOF` |
| 0.8 | `Cell.Input`: a cell as typed in Excel |

## Now: 0.9, building workbooks

The features people reach for EPPlus or ClosedXML to get, so a report can be built in code and
saved without a template.

- Filling a sheet from a collection or a `DataTable`, columns chosen by lambdas (no reflection, so
  NativeAOT keeps working).
- Hyperlinks on cells: web addresses, e-mail and places in the workbook.
- Data validation: lists, numbers, dates, text length and custom formulas, with input and error
  messages; read, changed and saved, and moved with inserted and deleted rows.
- Sheet and workbook protection with a password, and locked and hidden cells in cell styles.

## Next: 0.10, rules on cells

- Conditional formatting: read, added and saved. SharpCell can also tell which cells a rule
  applies to, since it calculates the formulas Excel would.
- Comments (notes): read, added and saved.
- Page setup: orientation, margins, print area, titles and page breaks.
- Benchmarks against EPPlus and ClosedXML: loading, recalculating and saving large workbooks.

## Later

- Cells that break their data validation, found by calculating the rules (Excel's Circle Invalid Data).
- Sorting a range, with formulas moving as Excel moves them.
- Fitting column widths to their contents. This needs font metrics, which the engine does not
  carry; it may come as a separate package.
- Resizing and renaming tables, and inserting columns into them.
- Removing and renaming sheets.
- Rich text in cells.

## Not planned

Creating charts, pivot tables, pictures or shapes (a file's own are kept when it is saved),
running macros, .xls and .xlsb files, and encrypted files.
