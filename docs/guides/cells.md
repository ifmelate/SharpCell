# Cells, ranges and dates

Going through the cells of a sheet, reading and writing blocks of values, dates, and the defined
names of a workbook.

## Going through cells

```csharp snippet=cells-enumerate
foreach (Cell cell in sheet.Cells)   // non-empty cells, row by row
{
    if (cell.Formula is not null)
        lines.Add($"{cell.Address}: {cell.Formula} = {cell.Value}");
}

CellRange? used = sheet.UsedRange;   // A1:F6 for the sample budget; null for an empty sheet
```

`Worksheet.Cells` visits only cells that hold something: a constant, a formula, or a value a
dynamic array spilled into them. Empty cells cost nothing, so a sheet with a cell in `XFD1048576`
is as quick to walk as any other. `CellRange.Cells` does the same for part of a sheet.

While enumerating you may change the value or formula of the cells you get. Adding or removing a
cell, including setting a value to `CellValue.Empty` or a recalculation that changes a spill,
makes the next step throw `InvalidOperationException`, as changing a .NET collection during a
`foreach` does.

## Ranges

`Worksheet.Range` takes an A1 address on that sheet: `A1:C10`, a single cell `B2`, whole columns
`A:B` or whole rows `2:3`. A `CellRange` is a handle, like a `Cell`: making one allocates nothing.
Its `ToString()` includes the sheet, as in `'My Sheet'!A1:C10`.

```csharp snippet=cells-values
CellRange amounts = sheet.Range("B2:B5");
CellValue[,] before = amounts.GetValues();   // [row, column] from B2: 1200, 450, 120, 40

amounts.SetValues(new CellValue[,] { { 1300 }, { 500 }, { 150 }, { 50 } });
workbook.Recalculate();
// sheet["B6"].Value is now 2000
```

`GetValues` returns every cell of the range, empty ones as `CellValue.Empty`, indexed
`[row, column]` from the top-left cell. A whole column has 1,048,576 cells and the array has as
many elements; for sparse data use `UsedRange` or `Cells`.

`SetValues` writes the way assigning `Cell.Value` does: formulas are replaced and
`CellValue.Empty` clears a cell. It checks everything first (the array's size, values a cell
cannot hold, cells inside an array formula other than its top-left cell, even when the range
covers the whole array) and writes nothing when a check fails. As with any change, formulas see the new values after `Workbook.Recalculate`.

## Dates

Excel stores a date as a number: days counted from the workbook's date system, with the time of
day as the fraction. A `CellValue` does not know which date system its workbook uses, so the
conversions take it:

```csharp snippet=cells-dates
sheet["A1"].Value = CellValue.DateTime(new DateTime(2026, 3, 15), workbook.DateSystem);
sheet["A2"].Formula = "=EOMONTH(A1, 0)";
workbook.Recalculate();

DateTime monthEnd = sheet["A2"].Value.AsDateTime(workbook.DateSystem);   // 2026-03-31
```

`AsDateTime` throws `InvalidOperationException` for a value that is not a number, like the other
`As` methods, and `ArgumentOutOfRangeException` for a number that is no date: a negative one, one
past 9999-12-31, or 60 in the 1900 date system, which Excel shows as the non-existent 29 February
1900. The time is rounded to the millisecond. `CellValue.DateTime` refuses dates before the date
system starts.

## Defined names

```csharp snippet=cells-names
foreach (DefinedName name in workbook.DefinedNames)
    lines.Add($"{name.Name} {name.Formula} {name.Scope?.Name ?? "(workbook)"}");   // TaxRate =0.2 (workbook)
```

`Workbook.DefinedNames` lists workbook-wide and sheet-scoped names in the order they were first
defined, spelled as they were defined. It is a snapshot: names defined later are not in it.
