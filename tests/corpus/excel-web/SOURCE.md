# Excel for the web corpus

Workbooks calculated by Microsoft Excel for the web on 2026-10-09, for behaviour the corpus in
`../excel` does not pin down: tables, hidden rows and filters.

Each input workbook was written without saved formula results and with `fullCalcOnLoad="1"`
by the script of the same name in `inputs/` (`python3 inputs/<name>.py <name>.xlsx`), uploaded
to OneDrive, opened in Excel for the web and saved with File > Save As > Download a Copy. The
downloaded copy holds the results Excel calculated. (Downloading from the OneDrive file list
returns the uploaded file unchanged, without results.)

| File | What it pins down |
|---|---|
| `tables-spike.xlsx` | SUBTOTAL and AGGREGATE options 0-7 on a filtered table and on a table with a row hidden by hand; `#This Row` in the header and totals rows, above and below a table |
| `hidden-rows.xlsx` | a row hidden by hand on sheets without a filter: a plain range, a sheet autoFilter without criteria, tables with and without a header row and filter buttons |
| `spill-blocked.xlsx` | a spill reference (`A3#`) to an anchor whose spill a typed value blocks is `#SPILL!`, not `#REF!` (see below) |
| `filter-mode.xlsx` | a sheet with a table filter, a sheet filter, a filter that hides nothing, and no filter: on a sheet with any filter criteria, every hidden row counts as filtered |
| `number-formats.xlsx` | `TEXT(value, code)` next to the same value shown through a cell style with that code: `Cell.Text` of column C must equal Excel's saved `TEXT` result in column B (Excel saves no display text) |
| `protection.xlsx` | the SHA-512 sheet protection hash Excel writes for the password `sharpcell` (see below) |

The same check run on `../excel/tables.xlsx` with its results removed reproduced all 212
results saved by desktop Excel.

`spill-blocked.xlsx` was made differently, on 2026-10-10: `tools/excel-oracle/writer-check/spill-blocked.xlsx`
(`excel/DynamicArrays.xlsx` with `B5` set to 999, written by SharpCell's `XlsxWriter`) opened in
Excel for the web and saved with Download a Copy. Excel recalculated it on opening: the `#REF!`
SharpCell had written for `TAKE(A3#,3,1)` and `SUM(A3#)` came back as `#SPILL!`.

`writer-*.xlsx` were made on 2026-10-10: SharpCell's `XlsxWriter` saved a workbook built in code
(`writer-new.xlsx`) and Excel's templates from `../excel/templates` with two rows inserted at row 4
of every sheet, a styled cell and a new sheet; `calcPr fullCalcOnLoad` was set, and each file was
opened in Excel for the web, recalculated, and saved with Download a Copy. Excel opened every file
without repairing it, and every formula and value matched what SharpCell saved, apart from `NOW()`.
`XlsxWriterExcelWebTests` builds the same workbooks again and compares them with these copies.
Excel for the web rewrote the theme's cached system colour and rounded some column widths and row
heights; neither is compared.

`protection.xlsx` was made on 2026-10-11 by protecting the sheet of an empty workbook in Excel for the
web with the password `sharpcell` (Review > Manage Protection, default options) and saving it with
Download a Copy. `XlsxWriterProtectionTests` checks that SharpCell verifies the password against
the `hashValue`, `saltValue` and `spinCount` Excel wrote.
