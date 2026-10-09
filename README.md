# SharpCell

A free (MIT) Excel formula engine for .NET: parser, dependency graph, recalculation,
dynamic arrays, LET/LAMBDA. Correctness is checked against workbooks calculated by real Excel.

Status: early development, nothing is published yet.

## Usage

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

```csharp snippet=build-in-code
var workbook = new Workbook();
Worksheet sheet = workbook.AddSheet("Sheet1");
sheet["A1"].Value = 2;
sheet["A2"].Formula = "=A1*3";
workbook.Recalculate();
CellValue result = sheet["A2"].Value;   // 6
```

A runnable version is in [samples/SharpCell.Sample](samples/SharpCell.Sample). `SharpCell.Xlsx` reads `.xlsx` files (not `.xls`, `.xlsm` or `.xlsb`); SharpCell does not write files.

## Documentation

- [Documentation site](https://ifmelate.github.io/SharpCell/): getting started, guides, API reference ([source](docs/index.md)).
- [What matches Excel](docs/compatibility.md), also as [JSON](docs/compatibility.json).
- [For AI coding agents](docs/agents.md): `llms.txt`, `llms-full.txt` and rules to paste into your agent's instructions.
