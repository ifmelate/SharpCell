# Sheets, filters and sorting

Renaming, removing, moving and hiding sheets; Excel's filter; sorting a range. Formulas follow each
change as they do in Excel.

## Managing sheets

```csharp snippet=sheets-manage
data.Name = "Data 2025";             // Summary!A1 is now ='Data 2025'!A1*2
workbook.MoveSheet(data, 0);         // Data 2025 comes first
workbook.RemoveSheet(draft);         // formulas reading Draft get #REF!
data.Visibility = SheetVisibility.Hidden;
summary.TabColor = CellColor.FromRgb(0x2E75B6);
workbook.Recalculate();
```

Setting `Worksheet.Name` renames the sheet in every formula that refers to it: in cells of every
sheet, defined names, data validation and conditional formats, quoting the name when it needs it.
Text inside `INDIRECT` and the targets of hyperlinks keep the old name, as in Excel. A name another
sheet has (ignoring case), or one Excel does not allow (more than 31 characters, `: \ / ? * [ ]`),
throws `ArgumentException`.

`Workbook.RemoveSheet` removes a sheet as Excel's Delete Sheet does. References to it become
`#REF!`: `=Draft!A1` becomes `=#REF!A1`. A reference through several sheets
(`=SUM(Jan:Mar!B2)`) whose first or last sheet is removed keeps the sheets left between its ends.
Names local to the sheet and its tables go with it, and references to those tables become `#REF!`.
The removed `Worksheet` can still be read; changing it throws `InvalidOperationException`
(`IsRemoved` is true).

`Workbook.MoveSheet` puts a sheet at another place in `Workbook.Sheets`. No formula changes, but a
reference through several sheets takes in the sheets now between its ends.

`Visibility` hides a sheet (`Hidden`, which Excel's Unhide lists) or hides it so that only code can
show it again (`VeryHidden`). A workbook keeps at least one visible sheet: hiding or removing the
last one throws `InvalidOperationException`. `TabColor` colours the sheet's tab.

Each change goes through every formula of the workbook, and every formula is calculated again at
the next `Recalculate`. When the workbook is saved, the file's sheets are renamed, reordered and
hidden, a removed sheet's part goes with its drawings, comments and tables, and charts, pivot table
sources, print areas and names in the file follow the new names (see [Saving .xlsx files](saving.md)).
A formula SharpCell cannot read (such as a link to another workbook) that names the sheet makes
renaming or removing it throw `NotSupportedException`.

## Filtering rows

```csharp snippet=sheets-filter
AutoFilter filter = sheet.SetAutoFilter("A1:B5");
filter.Filter(2, ColumnFilter.Custom(FilterOperator.GreaterThan, "100"));   // column B
filter.Filter(1, ColumnFilter.OneOf("North", "West"));                      // column A
bool hidden = sheet.IsRowHidden(3);   // true: South is not North or West

sheet["B2"].Value = 90;
filter.Apply();                       // Excel's Reapply: North is hidden now
```

`Worksheet.SetAutoFilter` turns on Excel's Data > Filter for a range whose first row holds the
buttons; `Worksheet.AutoFilter` is the filter a sheet has, read from a file too. `AutoFilter.Filter`
sets a column's criteria and applies the filter at once, as choosing them in Excel's menu does: each
row below the first shows when it meets the criteria of every filtered column, and hides otherwise.
The criteria are those of Excel's menu:

| Factory | Shows rows whose cell |
|---|---|
| `ColumnFilter.OneOf("North", "1,250.00")` | shows one of the texts, as formatted, ignoring case |
| `ColumnFilter.OneOf(new FilterDate(DateGrouping.Month, 2026, 3))` | holds a date in one of the years, months or days |
| `ColumnFilter.Blank()` | is empty; `OneOf(...) with { Blanks = true }` adds empty cells |
| `ColumnFilter.Custom(FilterOperator.GreaterThan, "100")` | meets a condition, or two with `matchAll` (And) or not (Or) |
| `ColumnFilter.BeginsWith`, `EndsWith`, `Contains`, `DoesNotContain` | holds the text there, ignoring case |
| `ColumnFilter.Top(10)`, `Top(10, percent: true)`, `Bottom(...)` | holds one of the highest or lowest numbers |
| `ColumnFilter.Of(DynamicFilter.AboveAverage)` | is above (or below) the column's average |
| `ColumnFilter.Of(DynamicFilter.ThisMonth)` | holds a date in a period around today, or in a quarter or month of any year |

A custom condition compares a number with numbers and a text with texts; `Equal` and `NotEqual`
match the cell's text with `*` and `?` as wildcards. Formulas out of date in the range are
calculated first. Like Excel, a filter does not apply itself again when the data changes: call
`Apply` (Excel's Reapply). `Clear(column)` and `Clear()` take criteria off, and
`Worksheet.RemoveAutoFilter` turns the filter off and shows the rows it hid. While criteria are set,
`Worksheet.FilterMode` is true, so `SUBTOTAL` with codes 1 to 11 skips the hidden rows (see
[Tables and hidden rows](tables.md)).

The filter and its criteria are saved, and move with inserted and deleted rows and columns. A
file's filter by colour or icon is kept, but `Apply` throws `NotSupportedException` while one is
set. Tables have filters of their own, which these are not: a sheet's filter cannot overlap a table.

## Sorting

```csharp snippet=sheets-sort
sheet.Range("A2:B5").Sort(new SortKey(2, Descending: true));   // by units, highest first
```

`CellRange.Sort` sorts the rows of a range as Excel's Data > Sort does, by one key or several
(`SortKey` is a sheet column and a direction); rows equal by every key keep their order. Ascending,
numbers come first, then text, ignoring case, then `FALSE` and `TRUE`, then errors; descending
reverses that, and empty cells come last either way. `matchCase: true` puts lower case first among
texts that differ only in case, and `leftToRight: true` sorts the columns by rows.

The cells move with their values, formulas, number formats, styles, hyperlinks and notes. A moved
formula keeps its relative references, as when it is copied: `=B2*C2` sorted into row 5 is
`=B5*C5`. Formulas elsewhere that read the range are not changed, so they read whatever was sorted
into their cells. Data validation, conditional formats, row heights and hidden rows stay where they
are. A range with merged cells, part of an array formula or a spill throws
`InvalidOperationException`; nothing moves.
