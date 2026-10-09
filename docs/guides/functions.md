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

`COUNTIF`, `SUMIF`, `AVERAGEIF`, the `*IFS` functions, the `D` database functions and `MATCH` take criteria written as in Excel:

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

## Dates and times

Dates and times are numbers: serial days, with the time of day as the fraction. Text that looks like a date or time, such as `"2024-01-10"` or `"12:00"`, is converted to a number when arithmetic or a function needs one, using `Workbook.Culture`. Two-digit years 00–29 mean 2000–2029.

## Regular expressions

`REGEXTEST`, `REGEXEXTRACT`, `REGEXREPLACE` and match mode 3 of `XLOOKUP` and `XMATCH` use .NET regular expressions, not PCRE2 as Excel does. Most patterns behave the same. `\d` and `\w` also match non-ASCII digits and letters, and PCRE-only syntax such as possessive quantifiers is not available.
