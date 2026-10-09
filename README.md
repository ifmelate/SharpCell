# SharpCell

A free (MIT) Excel formula engine for .NET: parser, dependency graph, recalculation,
dynamic arrays, LET/LAMBDA. Correctness is checked against workbooks calculated by real Excel.

Status: early development, nothing is published yet.

## Usage

```csharp
using SharpCell;
using SharpCell.Xlsx;

var workbook = XlsxReader.Load("report.xlsx");  // values and Excel's cached results
workbook["Sheet1"]["B2"].Value = 42;
workbook.Recalculate();
var total = workbook["Sheet1"]["B10"].Value;

var sum = new Workbook().Evaluate("=SUM(1,2,3)");  // formulas without a file
```

`SharpCell.Xlsx` reads `.xlsx` files (not `.xls`, `.xlsm` or `.xlsb`); SharpCell does not write files.
What matches Excel so far: [docs/compatibility.md](docs/compatibility.md).
