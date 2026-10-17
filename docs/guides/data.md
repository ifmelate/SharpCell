# Data in and out

Filling sheets from collections, `DataTable`s and CSV text, and reading ranges back out.

## Filling a sheet

```csharp snippet=data-in
sheet.LoadRows("A1", orders,
    ("Id", o => o.Id), ("Total", o => o.Total), ("Placed", o => o.Placed));

sheet.LoadCsv("E1", new StringReader("Code,Share\nX1,12%\n"));   // 12% is the number 0.12
```

`Worksheet.LoadRows` writes a header row and one row per item, its columns given as lambdas, so no
reflection is needed and the code trims and compiles ahead of time. `LoadDataTable` writes a
`DataTable`, with or without a header row. Strings stay text, never formulas; numbers, logical
values, dates and times become those values, dates with a date format. Nothing is written unless
every value could be stored.

`Worksheet.LoadCsv` reads CSV text as Excel opens a .csv file: each field as if typed into its cell
(see [Typing into cells](cells.md#typing-into-cells)), so numbers, percentages and dates in the
workbook's culture become those values with a matching format. Fields may be quoted, with doubled
quotes and line breaks inside. A field starting with `=` stays text unless `formulas: true`, since
a file from elsewhere could hide formulas in it. Pass `delimiter: ';'` where the decimal separator is
a comma, as Excel does there.

## Reading a range out

```csharp snippet=data-out
DataTable table = sheet.Range("A1:B5").ToDataTable();   // Region (string), Units (double)

List<(string Region, double Units)> rows = sheet.Range("A1:B5")
    .ReadRows(row => (row["Region"].Text, row["Units"].Value.AsNumber()));

string csv = sheet.Range("A1:B5").ToCsv();               // "Region,Units\r\nNorth,120\r\n..."
sheet.Range("A1:B5").WriteCsv(writer);                   // to a file or a stream
```

`CellRange.ToDataTable` gives a `DataTable` named after the sheet. The first row names the columns,
as tables name them: `Column3` for an empty header, `Units2` for a repeated one; pass
`headers: false` for `Column1`, `Column2` and so on. A column whose values are all numbers is
`double`, all shown as dates `DateTime`, all logical values `bool`, all text `string`, and any other
mix `object`, with errors as their text (`#N/A`). Empty cells are `DBNull`.

`CellRange.ReadRows` makes an object of each row below the first, with a lambda that picks cells by
header (ignoring case) or by sheet column. `ToCsv` and `WriteCsv` write each cell as it shows, with
its number format, as Excel's Save As CSV does, and end each row in CRLF.

Whole columns such as `A:C` end at the last row used. Formulas out of date in the range are
calculated first, so what is read out is current.
