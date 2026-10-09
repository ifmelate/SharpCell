# Recalculation

When cell values update, what is recalculated, and how to cancel, inspect problems, use threads and set the culture.

## Values update on Recalculate

```csharp snippet=recalculate-stale
sheet["A1"].Value = 5;
// sheet["A2"].Value is still 6: formulas are updated by Recalculate, not on assignment
workbook.Recalculate();
// sheet["A2"].Value is now 15
```

`Workbook.Recalculate` calculates the formulas that are out of date: those whose inputs changed since the last calculation, everything that depends on them, and formulas that use a volatile function. It does not recalculate the rest.

Volatile functions are recalculated on every `Recalculate`:

- `AGGREGATE`, `SUBTOTAL`
- `INDIRECT`, `OFFSET`
- `NOW`, `TODAY`
- `RAND`, `RANDARRAY`, `RANDBETWEEN`

`Workbook.Evaluate` calculates the out-of-date cells its formula reads before evaluating it.

## Cancellation

```csharp snippet=cancel
using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
try
{
    workbook.Recalculate(timeout.Token);
}
catch (OperationCanceledException)
{
    // The workbook stays usable: the next Recalculate finishes the work.
    cancelled = true;
}
```

Cells finished before the cancellation keep their new values. The rest stay out of date and are calculated by the next `Recalculate`.

## Diagnostics

Problems that do not show in the values alone are listed in `Workbook.Diagnostics`:

```csharp snippet=diagnostics
sheet["F1"].Formula = "=G1";
sheet["G1"].Formula = "=F1";
workbook.Recalculate();
foreach (CalculationDiagnostic diagnostic in workbook.Diagnostics)
{
    // CircularReference at Sheet1!F1: Circular reference: Sheet1!F1 -> Sheet1!G1 -> Sheet1!F1
    lines.Add($"{diagnostic.Kind} at {diagnostic.Sheet?.Name}!{diagnostic.Address}: {diagnostic.Message}");
}
```

| Kind | What happened | What the cell gets |
|---|---|---|
| `CircularReference` | Cells reference each other in a loop. | 0 in every cell of the loop, as Excel without iterative calculation |
| `FunctionFailure` | A function failed unexpectedly. This is a SharpCell bug worth reporting. | `#VALUE!` |
| `UnsupportedFormula` | A formula uses something SharpCell cannot evaluate, such as a table reference or a link to another workbook. | `#NAME?` |
| `LimitExceeded` | A result was larger than a limit allows, such as the cells all spills of a workbook may cover together. | An error |

Diagnostics describe the current state: editing a cell or calculating it without the problem removes its entries.

## Threads

```csharp snippet=threads
// A Workbook is not thread-safe. Separate workbooks are independent and can run in parallel.
Parallel.ForEach(paths, path =>
{
    Workbook workbook = XlsxReader.Load(path);
    workbook.Recalculate();
    totals.Add(workbook["Budget"]["B6"].Value.AsNumber());
});
```

## Culture and dates

```csharp snippet=culture
var german = new Workbook { Culture = CultureInfo.GetCultureInfo("de-DE") };
CellValue number = german.Evaluate("=VALUE(\"1,5\")");   // 1.5
```

`Workbook.Culture` controls text-to-number conversion, number-to-text conversion and text ordering. It defaults to the invariant culture, which behaves like English Excel. The process culture is never read.

`Workbook.DateSystem` says whether serial dates count from 1900 or from 1904; the .xlsx reader sets it from the file. In the 1900 system, serial number 60 is the non-existent 29 February 1900, as in Excel.
