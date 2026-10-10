# Files written by SharpCell, to check in Excel

Written by `tests/SharpCell.Tests/Xlsx/WriterCheckFiles.cs`. For each file: open it in Excel (or Excel for the web) and note whether it opens without repair, whether the cells below show these values before recalculating, and whether Formulas → Calculate Workbook changes anything.

## inputs-and-text.xlsx

From `../../samples/SharpCell.Sample/budget.xlsx`. Rent B2 changed to 1350; text constants in G2:G5 (escape-like text, CR LF, a control character, leading spaces); a new row 20.

- `Budget!B2`: 1350
- `Budget!B6`: 1960
- `Budget!F2`: 1040
- `Budget!G2`: "_x0041_"
- `Budget!G3`: "line\r\nbreak"
- `Budget!G4`: "\u0001 control"
- `Budget!G5`: "  leading spaces"
- `Budget!A20`: 42

## spill-grow.xlsx

From `excel/templates/invoice.xlsx`. A new item in the first empty cell of C15:C25: the item numbers in column B (SEQUENCE over COUNTA) grow by one.

- `B15`: 1
- `B16`: 2
- `B17`: 3
- `B18`: 4
- `B19`: <empty>
- `C18`: "Extra item"

## spill-shrink.xlsx

From `excel/templates/invoice.xlsx`. The last item of C15:C25 cleared: the item numbers in column B shrink by one.

- `B15`: 1
- `B16`: 2
- `B17`: <empty>
- `B18`: <empty>
- `C17`: <empty>

## volatile-spill-november.xlsx

From `excel/templates/weekly_timesheet.xlsx`. TODAY() fixed at 2026-11-15: the month's days in A3 (F1+SEQUENCE(DAY(EOMONTH(F1,0)))-1) shrink from 31 to 30; the weekday names next to them are text in a spill, written as formula results (t="str"). Excel recalculates TODAY() when it opens the file, so it will show its own month: check that the file opens without repair and that the weekday names are plain text.

- `Timesheet!F1`: 46341
- `Timesheet!A3`: 46341
- `Timesheet!B3`: "Sun"
- `Timesheet!A32`: 46370
- `Timesheet!B32`: "Mon"
- `Timesheet!A33`: <empty>
- `Timesheet!B33`: <empty>

## spill-blocked.xlsx

From `excel/DynamicArrays.xlsx`. B5 typed over the SEQUENCE spill at A3:B12: A3 is #SPILL! (a blocked spill, new rich value, the file had no xl/richData); F3 and K3 read it.

- `DynamicArrays!A3`: #SPILL!
- `DynamicArrays!F3`: #SPILL!
- `DynamicArrays!K3`: 0
- `DynamicArrays!B5`: 999

## rich-errors-cleared.xlsx

From `excel/calc_tests/INFORMATION/ISREF.xlsx`. M1 set to TRUE: D7 (was #CALC!) and D9 (was #SPILL!) get plain values and lose their vm; D11 shrinks to one cell.

- `D7`: 10
- `D9`: 10
- `D11`: 10
- `D12`: <empty>
- `D14`: 10

## keep-uncalculated.xlsx

From `excel/calc_tests/LOGICAL/IFERROR.xlsx`. Saved with KeepUncalculated: formulas calling functions SharpCell does not know keep Excel's results; the file asks Excel to recalculate on open.

