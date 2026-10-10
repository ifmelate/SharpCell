# Excel oracle

Builds SharpCell's own reference workbooks for behaviour the corpus in `tests/corpus/excel` does not cover
(LET/LAMBDA, the four kinds of "empty", spills, legacy and array formulas, dates, criteria).
Only Excel's own results count as reference values, so Excel calculates them.

## Run

On Windows with desktop Excel (Microsoft 365 or 2021+):

```powershell
powershell -ExecutionPolicy Bypass -File tools\excel-oracle\Build-Oracle.ps1
```

Each `cases/<name>.csv` becomes `tests/corpus/sharpcell/<name>.xlsx`; `<name>.1904.csv` uses the
1904 date system. `oracle-meta.json` records the Excel version and separators. Re-running
replaces the files. Then, on any machine:

```sh
dotnet run --project tools/SharpCell.Conformance   # updates docs/compatibility.md and .json
dotnet test
```

## Case format

`sheet,cell,content,kind` with a header row; see the comment at the top of `Build-Oracle.ps1`.
A formula is written in English with `,` separators, as in the formula bar of an English Excel.
A constant is a number with `.` decimals, `TRUE`/`FALSE`, or text; start text with `'` when it
would otherwise read as a number or be empty (`'` alone is the empty text `""`).

Add cases to an existing file by topic, or a new file for a new topic. A case without an Excel
result is not a reference: never edit the generated `.xlsx` files by hand.
