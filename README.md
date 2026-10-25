# SharpCell

[![NuGet](https://img.shields.io/nuget/v/SharpCell.Xlsx.svg)](https://www.nuget.org/packages/SharpCell.Xlsx)
[![ci](https://github.com/ifmelate/SharpCell/actions/workflows/ci.yml/badge.svg)](https://github.com/ifmelate/SharpCell/actions/workflows/ci.yml)

A free (MIT) Excel formula engine for .NET 8 and later: parser, dependency graph, recalculation,
dynamic arrays, LET/LAMBDA, tables with structured references (`Sales[Units]`) and hidden rows.
Correctness is checked against workbooks calculated by real Excel.

Status: 0.x. The API may change between minor versions until 1.0.

## Install

```bash
dotnet add package SharpCell.Xlsx   # reads and saves .xlsx; brings SharpCell with it
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

A runnable version is in [samples/SharpCell.Sample](samples/SharpCell.Sample). `SharpCell.Xlsx` reads `.xlsx` and `.xlsm` files (not `.xls` or `.xlsb`) saves them with your changes, and creates new ones, with data from objects or a `DataTable`, hyperlinks, data validation and protection ([Saving .xlsx files](docs/guides/saving.md)).

## Speed

Measured with BenchmarkDotNet on an Apple M4, .NET 10, on a generated sales model: a data sheet
with three formulas on every row (arithmetic, `IF`, a running total) and a summary sheet with
100 `SUMIFS` and 200 `XLOOKUP` over the whole data range, plus `SORT(UNIQUE(...))` and `LET`.

| Workbook | Load .xlsx | Calculate every formula | Change one input and recalculate | Save .xlsx |
|---|---:|---:|---:|---:|
| 10,000 rows: 90,000 cells, 30,000 formulas | 103 ms | 106 ms | 29 ms | 66 ms |
| 100,000 rows: 900,000 cells, 300,000 formulas | 861 ms | 1.8 s | 1.2 s | 683 ms |

Changing one input recalculates only what depends on it and allocates about 1.5 MB, whatever the
size of the workbook. The benchmarks are in [tools/SharpCell.Benchmarks](tools/SharpCell.Benchmarks):
`dotnet run --project tools/SharpCell.Benchmarks -c Release -- --filter "*"`.

## Documentation

- [Documentation site](https://ifmelate.github.io/SharpCell/): getting started, guides, API reference ([source](docs/index.md)).
- [Sheets, filters and sorting](docs/guides/sheets.md): renaming, removing and moving sheets with formulas following, Excel's filter, sorting a range.
- [Data in and out](docs/guides/data.md): sheets from collections, `DataTable`s and CSV, and ranges back out as `DataTable`s, objects and CSV.
- [Tables and hidden rows](docs/guides/tables.md): structured references, tables from files and code, SUBTOTAL and filters.
- [Custom functions](docs/guides/custom-functions.md): functions written in C# that formulas call, in place of VBA or add-ins.
- [Recalculation](docs/guides/recalculation.md): changed cells, precedents and dependents, cloning a template per request.
- [Roadmap](docs/roadmap.md): what was released and what comes next.
- [What matches Excel](docs/compatibility.md), also as [JSON](docs/compatibility.json).
- [For AI coding agents](docs/agents.md): `llms.txt`, `llms-full.txt` and rules to paste into your agent's instructions.

## Support

SharpCell is maintained by one person in spare time, with no support obligations. A bug report
needs a small .xlsx that shows the problem and the value Excel calculates; an issue without one
is closed. A `Workbook` is not thread-safe: use one per thread; separate workbooks are independent.
Security issues: see [SECURITY.md](SECURITY.md). Contributing: see [CONTRIBUTING.md](CONTRIBUTING.md).
