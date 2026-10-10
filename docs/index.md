# SharpCell

SharpCell is a free (MIT) Excel formula engine for .NET 8 and later: it reads .xlsx files, recalculates them, saves the results back and evaluates formulas, including dynamic arrays and LET/LAMBDA.

## Correctness

Results are compared with workbooks calculated by Microsoft Excel. The [compatibility report](compatibility.md) lists every function, its status and how many Excel-calculated cells it matches.

## Packages

| Package | What it does |
|---|---|
| `SharpCell` | The engine: parser, values, recalculation and functions. No dependencies. |
| `SharpCell.Xlsx` | Reads .xlsx files into a SharpCell workbook and saves recalculated values back. |

Install from NuGet:

```bash
dotnet add package SharpCell.Xlsx
```

## A first example

```csharp snippet=build-in-code
var workbook = new Workbook();
Worksheet sheet = workbook.AddSheet("Sheet1");
sheet["A1"].Value = 2;
sheet["A2"].Formula = "=A1*3";
workbook.Recalculate();
CellValue result = sheet["A2"].Value;   // 6
```

## Next

- [Getting started](getting-started.md): install, read a file, recalculate, read values.
- [Formulas and values](guides/formulas.md): syntax, references, dynamic arrays, names and LAMBDA.
- [Recalculation](guides/recalculation.md): when values update, cancellation, diagnostics, threads.
- [Reading .xlsx files](guides/xlsx.md): what the reader loads and what it skips.
- [Saving .xlsx files](guides/saving.md): writing recalculated values back into a template.
- [Functions](guides/functions.md): function status, criteria, dates.
- [What SharpCell does not do](guides/limits.md): every known gap.
- [API reference](api/index.md): every public type and member.
- [For AI coding agents](agents.md): llms.txt and rules for your agent.
