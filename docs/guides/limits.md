# What SharpCell does not do

Every known gap in one list, so you can decide before you start.

- **Writing files.** `XlsxWriter` saves changed values into the file a workbook was read from (see [Saving .xlsx files](saving.md)). It cannot create a file from a workbook built in code, or save new formulas, sheets, names, tables or styles.
- **Other file formats.** .xls, .xlsm, .xlsb and encrypted files throw `NotSupportedException`.
- **Tables do not change.** Tables cannot be resized, renamed or removed, and editing a header cell does not rename its column.
- **Filters are not applied.** Rows are hidden as the file or your code left them; filter criteria are not evaluated.
- **Iterative calculation.** A circular reference gives 0 in its cells and a diagnostic.
- **Display formats.** Cell number formats are not read. The `TEXT` function formats numbers itself.
- **Localized function names.** Formulas use English names, as in the file format.
- **Links to other workbooks.** They evaluate to `#NAME?`.
- **Cube, real-time data and web functions.** `CUBE*`, `RTD` and `WEBSERVICE` are not available.
- **Concurrent use.** One `Workbook` must not be used from several threads at once. Separate workbooks are independent.

Each function's status, and every function that differs from Excel in a described way, is in the [compatibility report](../compatibility.md).
