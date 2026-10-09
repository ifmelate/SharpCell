# SharpCell

A free (MIT) Excel formula engine for .NET 8 and later: it reads .xlsx files, recalculates them and evaluates formulas, including dynamic arrays and LET/LAMBDA. Results are checked against workbooks calculated by Microsoft Excel.

| Package | What it does |
|---|---|
| `SharpCell` | The engine: parser, values, recalculation and functions. No dependencies. |
| `SharpCell.Xlsx` | Reads .xlsx files into a SharpCell workbook. Brings `SharpCell` with it. |

## Read and recalculate a workbook

```csharp snippet=sample-usings
using SharpCell;
using SharpCell.Xlsx;
```

```csharp snippet=sample-load
// Read the workbook: sheets, values, formulas and any results Excel saved with them.
Workbook workbook = XlsxReader.Load(path);
Worksheet budget = workbook["Budget"];

// Calculate every formula now instead of relying on saved results.
workbook.Recalculate();
```

## Build a workbook in code

```csharp snippet=build-in-code
var workbook = new Workbook();
Worksheet sheet = workbook.AddSheet("Sheet1");
sheet["A1"].Value = 2;
sheet["A2"].Formula = "=A1*3";
workbook.Recalculate();
CellValue result = sheet["A2"].Value;   // 6
```

## Evaluate a formula without a cell

```csharp snippet=evaluate
CellValue sum = workbook.Evaluate("=SUM(1, 2, 3)");   // 6, no sheet needed
```

Values update only on `Recalculate`. Excel errors such as `#DIV/0!` are values, not exceptions. A `Workbook` is not thread-safe; use one per thread.

## Links

- [Documentation](https://ifmelate.github.io/SharpCell/)
- [What matches Excel](https://ifmelate.github.io/SharpCell/compatibility.html), function by function
- [What SharpCell does not do](https://ifmelate.github.io/SharpCell/guides/limits.html)
- [For AI coding agents](https://ifmelate.github.io/SharpCell/llms.txt)
- [Source and issues](https://github.com/ifmelate/SharpCell)
