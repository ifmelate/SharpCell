# Functions

Which functions exist, how criteria are written, and how dates and regular expressions behave.

## Function status

The [compatibility report](../compatibility.md) lists every function with one of three statuses, and the same data is in [compatibility.json](../compatibility.json):

| Status | Meaning |
|---|---|
| `implemented` | Matches Excel on the reference workbooks. |
| `known deviation` | Implemented, with a described difference from Excel. |
| `not implemented` | The function name evaluates to `#NAME?`. |

A function that does not appear in the report evaluates to `#NAME?` as well.

## Criteria

`COUNTIF`, `SUMIF`, `AVERAGEIF`, the `*IFS` functions and the `D` database functions take criteria written as in Excel:

| Criterion | Matches |
|---|---|
| `5` or `"=5"` | The number 5 and the text `"5"` |
| `">5"`, `"<=5"`, `"<>5"` | Numbers compared with 5. `<` and `>` compare only values of the same type. |
| `"<>"` | Any non-empty cell; the text `""` counts as non-empty |
| `""` | Empty cells and the text `""` |
| `"app*"`, `"?pple"` | Text with wildcards: `*` any run of characters, `?` one character |
| `"~*"` | A literal `*`; `~` escapes `*`, `?` and `~` |
| `TRUE`, `"true"` | Logical values only |

Wildcards apply to text only. In the `D` functions, a plain text criterion matches text that starts with it, as in Excel.

`MATCH` (match type 0), `XMATCH` and `XLOOKUP` (match mode 2) accept the same wildcards but none of the operators: `MATCH(">5", A1:A9, 0)` looks for the text `">5"`.

## Dates and times

Dates and times are numbers: serial days, with the time of day as the fraction. Text that looks like a date or time, such as `"2024-01-10"` or `"12:00"`, is converted to a number when arithmetic or a function needs one, using `Workbook.Culture`. Two-digit years 00–29 mean 2000–2029.

## Grouping

`GROUPBY` and `PIVOTBY` summarize a table the way a pivot table does, and `PERCENTOF` gives a
part of a total. The function argument is a function passed by name (`SUM`, `AVERAGE`,
`PERCENTOF`), a `LAMBDA`, or several of them side by side with `HSTACK`. A function of two
arguments, such as `PERCENTOF` or `LAMBDA(part, all, SUM(part)/SUM(all))`, also gets the values its
group is a part of (in `PIVOTBY`, the ones `relative_to` picks):

```csharp snippet=groupby
sheet["D1"].Formula = "=GROUPBY(A1:A3, B1:B3, SUM)";   // North 40, South 20, Total 60
sheet["G1"].Formula = "=GROUPBY(A1:A3, B1:B3, HSTACK(SUM, MAX))";   // a header row SUM, MAX, then one column each
workbook.Recalculate();
```

Groups are sorted ascending, empty keys last; totals are labelled `Total`, or `Grand Total` when
there are subtotals. The results were compared with Excel for the web cell by cell. In `PIVOTBY`,
several functions at once give `#VALUE!`.

## Regular expressions

`REGEXTEST`, `REGEXEXTRACT`, `REGEXREPLACE` and match mode 3 of `XLOOKUP` and `XMATCH` use .NET regular expressions, not PCRE2 as Excel does. Most patterns behave the same. `\d` and `\w` also match non-ASCII digits and letters, and PCRE-only syntax such as possessive quantifiers is not available.
