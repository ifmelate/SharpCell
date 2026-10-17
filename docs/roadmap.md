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
| 0.9 | Building workbooks: filling sheets from collections and `DataTable`s, hyperlinks, data validation, sheet and workbook protection, locked and hidden cells |
| 0.10 | Copying cells as Excel pastes them, formulas, links and data validation following (`CellRange.CopyTo`) |
| 0.11 | Conditional formats worked out as Excel shows them, notes, page setup, inserting and deleting cells with a shift |

## Now: 0.12, sheets and data

- Done: renaming, removing, moving and hiding sheets, with formulas following as in Excel, and tab
  colours.
- Done: Excel's filter: values, dates, custom conditions, top and bottom, above and below average
  and date periods, applied as Excel applies them, read and saved.
- Done: sorting a range by one key or several, cells moving with their formulas and formats.
- Done: reading ranges out as `DataTable`s, objects and CSV, and loading CSV.
- Benchmarks against EPPlus and ClosedXML: loading, recalculating and saving large workbooks.

## Next: 0.13, how sheets look

- Grouping rows and columns into outlines that collapse and expand.
- Rich text in cells: runs of different fonts in one cell.
- Styles of whole rows and columns, named cell styles, pattern and gradient fills, rotated and
  shrunk text, theme colours written as such.
- Headers and footers, paper size, scaling and fitting to pages.
- Fitting column widths to their contents. This needs font metrics, which the engine does not
  carry; it may come as a separate package.

## Later

- Resizing, renaming and removing tables, inserting columns into them, totals rows and table styles.
- Copying a sheet, within a workbook and to another.
- Pictures placed on a sheet.
- Cells that break their data validation, found by calculating the rules (Excel's Circle Invalid Data).
- Applying a table's own filter.

## Not planned

Creating charts, pivot tables or shapes (a file's own are kept when it is saved),
running macros, .xls and .xlsb files, and encrypted files.
