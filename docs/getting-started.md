# Getting started

Install SharpCell, read an .xlsx file, recalculate it and read the results, step by step through the console sample.

## Requirements

.NET 8 or .NET 10.

## Install

SharpCell is not on NuGet yet. Until the first release, clone the repository and add project references to `src/SharpCell` and `src/SharpCell.Xlsx`, as [the sample's project file](../samples/SharpCell.Sample/SharpCell.Sample.csproj) does.

After the first release:

```bash
dotnet add package SharpCell.Xlsx
```

`SharpCell.Xlsx` brings `SharpCell` with it. Use `dotnet add package SharpCell` alone if you build workbooks in code and never read files.

## The sample

[samples/SharpCell.Sample](../samples/SharpCell.Sample/README.md) is a console app built around `budget.xlsx`, one sheet named `Budget`:

| Cell | Content |
|---|---|
| A2:B5 | Rent 1200, Food 450, Transport 120, Internet 40 (monthly) |
| C2:C5 | Yearly amounts, `=B2*12` and so on |
| B6 | `=SUM(B2:B5)`, the monthly total |
| F1 | Income, 3000 |
| F2 | `=F1-B6`, what is left |

### Namespaces

```csharp snippet=sample-usings
using SharpCell;
using SharpCell.Xlsx;
```

### Load and recalculate

`XlsxReader.Load` reads sheets, values and formulas. Cells with formulas start out holding the results Excel saved in the file, if any. `Recalculate` calculates every formula with SharpCell and replaces those results.

```csharp snippet=sample-load
// Read the workbook: sheets, values, formulas and any results Excel saved with them.
Workbook workbook = XlsxReader.Load(path);
Worksheet budget = workbook["Budget"];

// Calculate every formula now instead of relying on saved results.
workbook.Recalculate();
```

### Read values

A cell's `Value` is a `CellValue`. Check its `Kind` before calling `AsNumber`, `AsText` or another `As` method: each one throws `InvalidOperationException` when the value is of another kind. Excel errors such as `#DIV/0!` are values of kind `Error`, not exceptions.

```csharp snippet=sample-read
CellValue total = budget["B6"].Value;   // =SUM(B2:B5)
output.WriteLine($"Total per month: {Format(total)}");
output.WriteLine($"Left after expenses: {Format(budget["F2"].Value)}");
```

The sample formats values with this helper:

```csharp snippet=sample-format
// A CellValue can be a number, text, logical, error, array or empty; check Kind before As*.
private static string Format(CellValue value) => value.Kind switch
{
    CellValueKind.Number => value.AsNumber().ToString(CultureInfo.InvariantCulture),
    CellValueKind.Error => value.AsError().ToText(),   // "#DIV/0!", "#N/A", ...
    _ => value.ToString(),
};
```

### Change an input

Assigning a value does not update the formulas that depend on it. Call `Recalculate` after a batch of changes: it calculates only the cells that are out of date.

```csharp snippet=sample-change
// Change an input. Dependent formulas are updated by the next Recalculate, not before.
budget["B2"].Value = 1350;
workbook.Recalculate();
output.WriteLine("After rent change to 1350:");
output.WriteLine($"Total per month: {Format(budget["B6"].Value)}");
output.WriteLine($"Left after expenses: {Format(budget["F2"].Value)}");
```

### Formulas without a cell

`Workbook.Evaluate` calculates a formula that lives in no cell. It can read the workbook's sheets. A formula that returns several values gives a `CellValue` of kind `Array`; its text form is `{1;2;3}`.

```csharp snippet=sample-evaluate
// Evaluate formulas that live in no cell. They can read the workbook's sheets.
CellValue food = workbook.Evaluate("=XLOOKUP(\"Food\", Budget!A2:A5, Budget!B2:B5)");
CellValue average = workbook.Evaluate("=LET(items, Budget!B2:B5, SUM(items) / COUNT(items))");
CellValue sequence = workbook.Evaluate("=SEQUENCE(3)");   // a dynamic array: {1;2;3}
output.WriteLine($"Food per month: {Format(food)}");
output.WriteLine($"Average expense: {Format(average)}");
output.WriteLine($"Sequence: {sequence}");
```

## Run it

```bash
dotnet run --project samples/SharpCell.Sample -f net10.0
```

```text
Total per month: 1810
Left after expenses: 1190
After rent change to 1350:
Total per month: 1960
Left after expenses: 1040
Food per month: 450
Average expense: 490
Sequence: {1;2;3}
```

## Next

- [Formulas and values](guides/formulas.md)
- [Recalculation](guides/recalculation.md)
- [Reading .xlsx files](guides/xlsx.md)
