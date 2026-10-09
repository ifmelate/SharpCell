# Formulas and values

How to write formulas, what a cell value can be, and how dynamic arrays, names and LAMBDA work.

## Syntax

Write formulas the way Excel stores them in a file: English function names, a comma between arguments and a dot as the decimal separator. The leading `=` is optional; `Cell.Formula` adds it if missing. Localized function names such as `СУММ` or `SOMME` are not recognized.

References:

| Form | Meaning |
|---|---|
| `A1`, `$A$1` | A cell, relative or absolute |
| `A1:C3`, `A:A`, `1:1` | A range, whole columns, whole rows |
| `'Sheet name'!A1` | A cell on another sheet |
| `Sheet1:Sheet3!A1` | The same cell on a run of sheets |
| `TaxRate` | A defined name |
| `C1#` | The whole range a dynamic array in C1 spilled into |
| `@A1:A9` | Implicit intersection: the cell of the range in the formula's row or column |

## Build a workbook in code

```csharp snippet=build-in-code
var workbook = new Workbook();
Worksheet sheet = workbook.AddSheet("Sheet1");
sheet["A1"].Value = 2;
sheet["A2"].Formula = "=A1*3";
workbook.Recalculate();
CellValue result = sheet["A2"].Value;   // 6
```

## Reading values

A `CellValue` is one of: `Number`, `Text`, `Boolean`, `Error`, `Array`, `Empty`, `Missing` or `Lambda`. Dates and times are numbers: serial days since the workbook's date system epoch. Errors such as `#DIV/0!` are values, not exceptions.

```csharp snippet=read-values
static string Describe(CellValue value) => value.Kind switch
{
    CellValueKind.Number => value.AsNumber().ToString(CultureInfo.InvariantCulture),
    CellValueKind.Text => value.AsText(),
    CellValueKind.Boolean => value.AsBoolean() ? "TRUE" : "FALSE",
    CellValueKind.Error => value.AsError().ToText(),   // "#DIV/0!", "#N/A", ...
    CellValueKind.Array => $"array of {value.AsArray().GetLength(0)} x {value.AsArray().GetLength(1)}",
    CellValueKind.Empty => "(empty)",
    _ => value.ToString(),
};
```

The `As` methods throw `InvalidOperationException` when the value is of another kind. Every member is listed in the [CellValue reference](../api/SharpCell.CellValue.md).

## Dynamic arrays

A formula that returns several values spills into the cells below and to the right. The spilled cells hold values but no formula: their `Formula` is `null`.

```csharp snippet=spill
sheet["C1"].Formula = "=SEQUENCE(3)";   // spills into C1:C3
sheet["D1"].Formula = "=SUM(C1#)";      // C1# is the whole spilled range
workbook.Recalculate();
// sheet["C3"].Value is 3 and sheet["D1"].Value is 6
```

When a cell in the way holds something, the anchor shows `#SPILL!` and the other cells stay as they are:

```csharp snippet=spill-blocked
sheet["C2"].Value = "in the way";
workbook.Recalculate();
// sheet["C1"].Value is #SPILL!; clear C2 and recalculate to spill again
```

## Names and LAMBDA

`DefineName` gives a formula a name. A name whose formula is a `LAMBDA` can be called like a function.

```csharp snippet=names-lambda
workbook.DefineName("TaxRate", "=0.2");
workbook.DefineName("WithTax", "=LAMBDA(amount, amount * (1 + TaxRate))");
CellValue gross = workbook.Evaluate("=WithTax(100)");   // 120
```

## Formulas outside cells

```csharp snippet=evaluate
CellValue sum = workbook.Evaluate("=SUM(1, 2, 3)");   // 6, no sheet needed
```

## Errors in formula text

Text that is not a valid formula throws `FormulaParseException`, which says where parsing stopped. The cell is left unchanged.

```csharp snippet=parse-error
string? problem = null;
try
{
    sheet["A1"].Formula = "=SUM(1,";
}
catch (FormulaParseException e)
{
    problem = $"at position {e.Position}: {e.Message}";
}
```

A formula read from a file that SharpCell cannot handle does not throw. Once calculated, it is `#NAME?` and listed in `Workbook.Diagnostics` as `UnsupportedFormula`. See [Reading .xlsx files](xlsx.md).
