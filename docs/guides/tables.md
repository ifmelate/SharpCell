# Tables and hidden rows

An Excel table is a named range with column names. Formulas refer to it with structured
references such as `Sales[Units]`; SharpCell evaluates them against the table as it is.

## Tables in code

```csharp snippet=tables-add
var workbook = new Workbook();
Worksheet sheet = workbook.AddSheet("Sales");
sheet["A1"].Value = "Region";
sheet["B1"].Value = "Units";
sheet["A2"].Value = "North";
sheet["B2"].Value = 10;
sheet["A3"].Value = "South";
sheet["B3"].Value = 20;
sheet.AddTable("Sales", "A1:B3");

sheet["D1"].Formula = "=SUM(Sales[Units])";
sheet["C2"].Formula = "=Sales[@Units]*2";   // the table's row on the formula's own row
workbook.Recalculate();
```

`AddTable` takes the column names from the header row as it is when the table is added; a header
formula counts by its value, as Excel turns it into text. A single cell, such as `"A1"`, is a
one-cell table without a header row. A table
does not change afterwards: SharpCell does not resize or rename tables, and editing a header cell
does not rename its column. Tables and defined names share one set of names.

## Structured references

| Reference | Cells |
|---|---|
| `Sales[Units]`, `Sales[[#Data],[Units]]` | the data rows of a column |
| `Sales`, `Sales[]`, `Sales[#Data]` | all data rows |
| `Sales[#All]` | the whole table, header and totals rows included |
| `Sales[#Headers]`, `Sales[#Totals]` | the header or totals row; `#REF!` if the table has none |
| `Sales[[#Headers],[#Data]]`, `Sales[[#Data],[#Totals]]` | two adjacent parts |
| `Sales[[Jan]:[Mar]]` | a range of columns |
| `Sales[@Units]`, `Sales[[#This Row],[Units]]` | the data row on the formula's own row; `#VALUE!` on any other row |
| `[Units]`, `[@Units]` | `Sales[Units]` and `Sales[@Units]`, written in a formula inside the table |

Inside a name, an apostrophe escapes `[`, `]`, `#` and `'`: the column `[est] Q1` is
`Sales['[est'] Q1]`. A missing table or column is `#REF!`. A formula inside a table that returns
several values is `#SPILL!`, as in Excel.

`Workbook.Evaluate` runs a formula that belongs to no cell, so it is in no table and has no row:
there `[Units]` is `#REF!` and `Sales[@Units]` is `#VALUE!`, while `Sales[Units]` works.

## Tables from .xlsx files

`XlsxReader` reads every table of the file: its name, range, column names and whether it has
header and totals rows. `Workbook.Tables` lists them.

## Hidden rows and filters

```csharp snippet=tables-hidden-rows
sheet.SetRowHidden(2, true);
sheet["B1"].Formula = "=SUBTOTAL(9,A1:A3)";     // counts a row hidden by hand
sheet["B2"].Formula = "=SUBTOTAL(109,A1:A3)";   // skips it
workbook.Recalculate();                         // B1 is 7, B2 is 5

sheet.FilterMode = true;                        // as if a filter hid the row
workbook.Recalculate();                         // both are 5
```

- `SUBTOTAL` with codes 101–111 and `AGGREGATE` with options 1, 3, 5 and 7 skip hidden rows.
- `SUBTOTAL` with codes 1–11 counts rows hidden by hand and skips rows a filter hid. A file does
  not say which is which, so Excel decides by the sheet: on a sheet with a filter that has
  criteria, every hidden row counts as filtered. `Worksheet.FilterMode` is that flag; the reader
  sets it from the file.
- `AGGREGATE` with options 0, 2, 4 and 6 counts every row.
- SharpCell does not apply filter criteria: hidden rows are as the file or your code left them.
