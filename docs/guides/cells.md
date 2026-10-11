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
does not block a spill and is not listed by `Cells`. `XlsxWriter` saves formats with the cells.

## Typing into cells

```csharp snippet=cells-input
sheet["A1"].Input = "12.5%";          // 0.125, and the format 0.00%
sheet["A2"].Input = "10/11/2026";     // a date in the workbook's culture, and its short date format
sheet["A3"].Input = "=A1*2";          // a formula
sheet["A4"].Input = "'0123";          // text, kept as typed
string edit = sheet["A1"].Input;       // "12.5%", as Excel shows it to edit
```

`Cell.Input` is what a person types into a cell in Excel. Setting it does what typing and
pressing Enter does: `=` starts a formula, `'` keeps the rest as text, `TRUE`, `FALSE` and error
literals are those values, and numbers, percentages, currency, grouped digits, dates and times are
read in `Workbook.Culture` by the same rules as text meeting arithmetic and `DATEVALUE`; anything
else stays text, and an empty string clears the cell. A cell in the General format takes a format
from what was typed: `12%` gives `0%`, `1,250.50` gives `#,##0.00`, `$5` a currency format, a date
the culture's short date, a time `h:mm` with seconds and AM/PM when they were typed. A cell with a
format of its own keeps it.

Getting `Input` gives the text Excel shows to edit the cell: the formula, a number in full, a
percentage as one, a date or time as the culture writes it, and text with an apostrophe in front
when it would read as something else. Typed back, it gives the same content. A format taken from
typing is saved with the cell by `XlsxWriter`. The formats Excel gives typed currency and dates are
not yet compared with Excel's.

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

`XlsxWriter` saves styles, sizes, merged cells, frozen panes and gridlines. A new style starts from
the cell's style in the file, so what SharpCell does not read (pattern fills, rotated text) stays.

## Inserting and deleting rows and columns

```csharp snippet=insert-rows
sheet.InsertRows(3);             // a new row 3; rows 3 and below move down
sheet["A3"].Value = "Travel";
sheet["B3"].Value = 300;
workbook.Recalculate();          // B7 is =SUM(B2:B6) now and includes the new row

sheet.DeleteColumns(5);          // column E goes; formulas that read it get #REF!
```

`InsertRows`, `DeleteRows`, `InsertColumns` and `DeleteColumns` work like Excel's commands. Cells,
styles, number formats, row heights, column widths, hidden rows and columns, merged areas and
tables move, and every reference to the moved cells follows them: in formulas of every sheet and
in defined names. A range the new rows are inserted into grows (`SUM(B2:B5)` becomes `SUM(B2:B6)`),
and so do merged areas and tables; inserting at a range's first row moves it instead. A range that
loses some of its rows shrinks, a reference to a deleted cell becomes `#REF!`, and a table whose
rows are all deleted is removed. Whole-column references such as `A:A` do not change when rows
move. References through several sheets (`Sheet1:Sheet3!A1`) and text given to `INDIRECT` do not
change either, as in Excel.

Every formula is calculated again at the next `Workbook.Recalculate`. An edit that would push
content past the last row or column, split an array formula (Ctrl+Shift+Enter), or remove a
table's header or totals row without the whole table throws `InvalidOperationException` and changes
nothing. Inserting or deleting columns inside a table throws `NotSupportedException`, and so does
any edit of a sheet that has, or is named by, a formula SharpCell cannot read (such as a link to
another workbook): its references could not follow. Formulas that name a table whose rows were all
deleted get `#REF!`. Each edit goes through every formula of the workbook, so insert or delete many
rows with one call (`InsertRows(5, 100)`) rather than one at a time. When the
workbook is saved, conditional formats, data validation, pictures, comments, charts and pivot table
sources in the file move with the cells too (see [Saving .xlsx files](saving.md)).

## Copying cells

```csharp snippet=cells-copy
sheet["B1"].Formula = "=A1*$A$1";
sheet.Range("B1").CopyTo(sheet.Range("B2:B3"));   // B2 is =A2*$A$1, B3 is =A3*$A$1
sheet.Range("A1:B3").CopyTo(sheet.Range("D1"));   // values, formulas, formats, styles, links, validation
workbook.Recalculate();
```

`CellRange.CopyTo` copies cells as copying and pasting them in Excel does: values, formulas,
number formats, styles, hyperlinks and data validation. Relative references in a formula keep their distance from its cell and
absolute ones (`$A$1`) stay; a reference that would point past the sheet's edge becomes `#REF!`.
Data validation formulas keep their distance from each cell in the same way. Empty cells of
the range clear the cells they land on, their validation and hyperlinks too. A destination that is a whole number of
times as tall and as wide as the range is filled by repeating it, so one cell copied to `B2:B3`
fills both; any other destination gets one copy at its top-left cell. The range and the
destination may overlap, and the destination may be in another workbook, where formulas refer to
that workbook's sheets and names. Merged cells, column widths and row heights are not copied. A
range holding part of an array formula (Ctrl+Shift+Enter) or a formula SharpCell cannot read throws
`NotSupportedException`; nothing changes.

## Defined names

```csharp snippet=cells-names
foreach (DefinedName name in workbook.DefinedNames)
    lines.Add($"{name.Name} {name.Formula} {name.Scope?.Name ?? "(workbook)"}");   // TaxRate =0.2 (workbook)
```

`Workbook.DefinedNames` lists workbook-wide and sheet-scoped names in the order they were first
defined, spelled as they were defined. It is a snapshot: names defined later are not in it.
