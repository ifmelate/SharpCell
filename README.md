# SharpCell

[![NuGet](https://img.shields.io/nuget/v/SharpCell.Xlsx.svg)](https://www.nuget.org/packages/SharpCell.Xlsx)
[![ci](https://github.com/ifmelate/SharpCell/actions/workflows/ci.yml/badge.svg)](https://github.com/ifmelate/SharpCell/actions/workflows/ci.yml)

A free (MIT) Excel formula engine for .NET 8 and later: parser, dependency graph, recalculation,
dynamic arrays, LET/LAMBDA, tables with structured references (`Sales[Units]`) and hidden rows.
Correctness is checked against workbooks calculated by real Excel.

Status: 0.x. The API may change between minor versions until 1.0.

## Install

```bash
dotnet add package SharpCell.Xlsx   # reads .xlsx; brings SharpCell with it
dotnet add package SharpCell        # the engine alone, for workbooks built in code
```

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
- [Tables and hidden rows](docs/guides/tables.md): structured references, tables from files and code, SUBTOTAL and filters.
- [What matches Excel](docs/compatibility.md), also as [JSON](docs/compatibility.json).
- [For AI coding agents](docs/agents.md): `llms.txt`, `llms-full.txt` and rules to paste into your agent's instructions.

## Support

SharpCell is maintained by one person in spare time, with no support obligations. A bug report
needs a small .xlsx that shows the problem and the value Excel calculates; an issue without one
is closed. A `Workbook` is not thread-safe: use one per thread; separate workbooks are independent.
Security issues: see [SECURITY.md](SECURITY.md). Contributing: see [CONTRIBUTING.md](CONTRIBUTING.md).
