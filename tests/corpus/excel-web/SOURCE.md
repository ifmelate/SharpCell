# Excel for the web corpus

Workbooks calculated by Microsoft Excel for the web on 2026-10-09, for behaviour the IronCalc
corpus does not pin down: tables, hidden rows and filters.

Each input workbook was written without saved formula results and with `fullCalcOnLoad="1"`
by the script of the same name in `inputs/` (`python3 inputs/<name>.py <name>.xlsx`), uploaded
to OneDrive, opened in Excel for the web and saved with File > Save As > Download a Copy. The
downloaded copy holds the results Excel calculated. (Downloading from the OneDrive file list
returns the uploaded file unchanged, without results.)

| File | What it pins down |
|---|---|
| `tables-spike.xlsx` | SUBTOTAL and AGGREGATE options 0-7 on a filtered table and on a table with a row hidden by hand; `#This Row` in the header and totals rows, above and below a table |
| `hidden-rows.xlsx` | a row hidden by hand on sheets without a filter: a plain range, a sheet autoFilter without criteria, tables with and without a header row and filter buttons |
| `filter-mode.xlsx` | a sheet with a table filter, a sheet filter, a filter that hides nothing, and no filter: on a sheet with any filter criteria, every hidden row counts as filtered |

The same check run on `../ironcalc/tables.xlsx` with its results removed reproduced all 212
results saved by desktop Excel.
