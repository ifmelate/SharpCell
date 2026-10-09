# What SharpCell does not do

Every known gap in one list, so you can decide before you start.

- **Writing files.** SharpCell reads .xlsx and keeps the workbook in memory; it does not save.
- **Other file formats.** .xls, .xlsm, .xlsb and encrypted files throw `NotSupportedException`.
- **Structured references.** Table references such as `Table1[Col]` evaluate to `#NAME?`.
- **Hidden rows.** Rows are never hidden or filtered, so `SUBTOTAL` and `AGGREGATE` include every row.
- **Iterative calculation.** A circular reference gives 0 in its cells and a diagnostic.
- **Display formats.** Cell number formats are not read. The `TEXT` function formats numbers itself.
- **Localized function names.** Formulas use English names, as in the file format.
- **Links to other workbooks.** They evaluate to `#NAME?`.
- **Cube, real-time data and web functions.** `CUBE*`, `RTD` and `WEBSERVICE` are not available.
- **Concurrent use.** One `Workbook` must not be used from several threads at once. Separate workbooks are independent.

Each function's status, and every function that differs from Excel in a described way, is in the [compatibility report](../compatibility.md).
