# Cells, ranges and dates

Going through the cells of a sheet, reading and writing blocks of values, dates, number formats
and the text Excel shows, and the defined names of a workbook.

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

## Number formats and display text

```csharp snippet=cells-text
sheet["A1"].Value = 1234.5;
sheet["A1"].NumberFormat = "#,##0.00";   // as read from the file, or set here
string shown = sheet["A1"].Text;          // "1,234.50"
```

`Cell.NumberFormat` is the cell's format code, such as `0.00`, `#,##0` or `yyyy-mm-dd`, and
`General` when it has none. `XlsxReader` reads it from the file's cell styles, Excel's built-in
formats included. `Cell.Text` is the value as Excel shows it through that format, with the
workbook's culture for separators and month names: dates come out as dates, percentages with
`%`, and `General` the way Excel's default column shows a number.

| Value | `Text` |
|---|---|
| Number | formatted by the cell's format; a date format on a number that is no date gives `#######` |
| Text | the format's text section (`"Name: "@`), or the text itself |
| Logical | `TRUE` or `FALSE` |
| Error | `#DIV/0!`, `#N/A` and so on |
| Empty | an empty string |

Column width plays no part, so a long number is never cut to `####`. A format is no content: it
does not block a spill and is not listed by `Cells`. Formats are not saved: `XlsxWriter` refuses a
workbook whose formats changed.

## Styles

```csharp snippet=cells-styles
CellStyle header = sheet["A1"].Style with   // the workbook's default until the cell has its own
{
    Font = sheet["A1"].Style.Font with { Bold = true },
    Fill = CellColor.FromRgb(0xD9E1F2),
    BottomBorder = new CellBorder(CellBorderStyle.Thin),
    HorizontalAlignment = CellHorizontalAlignment.Center,
};
sheet["A1"].Style = header;
bool bold = sheet["A1"].Style.Font.Bold;      // true
string font = sheet["B1"].Style.Font.Name;    // "Calibri", the default
```

`Cell.Style` is how the cell looks in Excel: its `CellFont` (name, size in points, bold,
italic, underline, strikethrough, colour), a fill colour, the four edges of its border and its
alignment, wrapping and indent. A cell without a style of its own has `Workbook.DefaultStyle`,
which `XlsxReader` takes from the file's Normal style; setting `Style` to null gives it back.
Styles are records: they compare by value, and `with` makes a changed copy.

`XlsxReader` turns theme and indexed colours into RGB with the file's theme and palette. Only
solid fills are read; patterns and gradients read as no fill. Styles of whole rows and columns
(for cells that hold nothing), conditional formats, rotated text, shrink to fit and diagonal
borders are not read. A style plays no part in calculation and is no content.

## Column widths, row heights and merged cells

```csharp snippet=cells-geometry
sheet.SetColumnWidth(2, 20);      // column B, in the units .xlsx files use
sheet.SetRowHeight(1, 30);        // row 1, in points
sheet.SetColumnHidden(3, true);   // column C
sheet.Merge("A1:D1");
double? width = sheet.ColumnWidth(1);           // null: column A has the default width
string merged = sheet.MergedAreas[0].Address;   // "A1:D1"
```

Column widths are in the units .xlsx files use: characters of the widest digit of the default
font, padding included, so Excel's standard 8.43 characters of Calibri 11 are 9.140625 here.
Row heights are in points. Null means the sheet's default, `Worksheet.DefaultColumnWidth` and
`Worksheet.DefaultRowHeight`, which are null too when the file gives none. Converting either to
pixels depends on the font and the screen and is left to the application.

Hidden columns change no result, as in Excel; hidden rows do (see
[Tables and hidden rows](tables.md)). `Merge` keeps the values of the cells it covers, where
Excel's Merge command would clear all but the top-left one.

## Frozen panes and gridlines

```csharp snippet=cells-view
sheet.FrozenRows = 1;          // the header row stays in place while the rest scrolls
sheet.FrozenColumns = 1;       // and so does column A
sheet.ShowGridlines = false;   // as View > Gridlines turned off in Excel
```

`FrozenRows` and `FrozenColumns` are the rows at the top and the columns on the left that Excel's
Freeze Panes keeps in place; `ShowGridlines` is whether Excel draws gridlines on the sheet.
`XlsxReader` takes them from the sheet's first view. A plain split that is not frozen reads as no
frozen panes. None of them changes a result.

Styles, sizes, merged cells, frozen panes and gridlines are not saved: `XlsxWriter` copies them
from the file and refuses a workbook in which they changed.

## Defined names

```csharp snippet=cells-names
foreach (DefinedName name in workbook.DefinedNames)
    lines.Add($"{name.Name} {name.Formula} {name.Scope?.Name ?? "(workbook)"}");   // TaxRate =0.2 (workbook)
```

`Workbook.DefinedNames` lists workbook-wide and sheet-scoped names in the order they were first
defined, spelled as they were defined. It is a snapshot: names defined later are not in it.
