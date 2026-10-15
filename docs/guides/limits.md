# What SharpCell does not do

Every known gap in one list, so you can decide before you start.

- **Writing files.** `XlsxWriter` saves what SharpCell models (see [Saving .xlsx files](saving.md)). It cannot add charts, pivot tables, threaded comments, pictures, colour scales, data bars or icon sets, or remove or rename sheets; a file's own are kept as they are.
- **Other file formats.** .xls, .xlsb and encrypted files throw `NotSupportedException`. Macros of .xlsm files are kept, never run.
- **Tables change only with rows.** Tables grow and shrink with inserted and deleted rows, but cannot be resized, renamed or removed directly, columns cannot be inserted into or deleted from them, and editing a header cell does not rename its column.
- **Filters are not applied.** Rows are hidden as the file or your code left them; filter criteria are not evaluated.
- **Iterative calculation is not compared with Excel.** It follows Excel's documented behaviour, but Excel for the web cannot calculate it, so no reference workbook pins the order of passes in a cycle of several cells.
- **Styles are partly read.** Fonts, solid fills, borders, alignment, number formats and cell protection (locked, hidden) are read and saved; pattern and gradient fills and rotated text are not read (a file's own are kept when it is saved). The colours of colour scales, data bars and icon sets are not worked out.
- **Protection is not enforced.** Sheet and workbook protection is read, set and saved, but SharpCell does not enforce it: code can still change a protected sheet or workbook.
- **Data validation is stored, not evaluated.** Rules are read, set and saved; nothing checks a value against them. Validations in `extLst` (`x14`) are not modelled, and neither is the `imeMode` attribute: it is lost when a sheet's rules are rewritten, which includes any row or column insert or delete on that sheet.
- **Large hyperlinks.** A hyperlink covering more than 10,000 cells, or more than 100,000 linked cells on a sheet, is kept in the file but not read into the model.
- **Hyperlink style.** Setting a hyperlink does not apply Excel's blue underlined Hyperlink style; style the cell yourself.
- **Localized function names.** Formulas use English names, as in the file format.
- **Links to other workbooks.** They evaluate to `#NAME?`.
- **Cube, real-time data and web functions.** `CUBE*`, `RTD` and `WEBSERVICE` are not available.
- **Concurrent use.** One `Workbook` must not be used from several threads at once. Separate workbooks are independent, and one workbook may be cloned from several threads while nothing changes it (see [Templates in a service](recalculation.md#templates-in-a-service)).

Each function's status, and every function that differs from Excel in a described way, is in the [compatibility report](../compatibility.md).
