# What SharpCell does not do

Every known gap in one list, so you can decide before you start.

- **Writing files.** `XlsxWriter` saves what SharpCell models (see [Saving .xlsx files](saving.md)). It cannot add charts, pivot tables, comments, pictures, conditional formats or data validation, or remove or rename sheets; a file's own are kept as they are.
- **Other file formats.** .xls, .xlsb and encrypted files throw `NotSupportedException`. Macros of .xlsm files are kept, never run.
- **Tables change only with rows.** Tables grow and shrink with inserted and deleted rows, but cannot be resized, renamed or removed directly, columns cannot be inserted into or deleted from them, and editing a header cell does not rename its column.
- **Filters are not applied.** Rows are hidden as the file or your code left them; filter criteria are not evaluated.
- **Iterative calculation is not compared with Excel.** It follows Excel's documented behaviour, but Excel for the web cannot calculate it, so no reference workbook pins the order of passes in a cycle of several cells.
- **Styles are partly read.** Fonts, solid fills, borders, alignment and number formats are read and saved; pattern and gradient fills, rotated text and conditional formats are not read (a file's own are kept when it is saved).
- **Localized function names.** Formulas use English names, as in the file format.
- **Links to other workbooks.** They evaluate to `#NAME?`.
- **Cube, real-time data and web functions.** `CUBE*`, `RTD` and `WEBSERVICE` are not available.
- **Concurrent use.** One `Workbook` must not be used from several threads at once. Separate workbooks are independent, and one workbook may be cloned from several threads while nothing changes it (see [Templates in a service](recalculation.md#templates-in-a-service)).

Each function's status, and every function that differs from Excel in a described way, is in the [compatibility report](../compatibility.md).
