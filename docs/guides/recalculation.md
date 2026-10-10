# Recalculation

When cell values update, what is recalculated, which cells changed, which cells a formula reads, and how to cancel, inspect problems, use threads and templates, and set the culture.

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

## Changed cells

`Recalculate` returns the cells whose values calculation changed, for example to send only those
to a client:

```csharp snippet=changed-cells
sheet["B2"].Value = 1350;
RecalculationResult result = workbook.Recalculate();
foreach (Cell cell in result.ChangedCells)   // C2, F2, F3, B6, C6: formulas whose value changed
    sent.Add($"{cell.Worksheet.Name}!{cell.Address} = {cell.Value}");
```

`ChangedCells` lists formula cells and cells a spill filled, changed or left (those are now empty),
in the order of the sheets, then by row and column. It covers everything calculated since the
previous `Recalculate`, including cells `Evaluate` calculated in between. Cells you changed
yourself are not in it, nor cells that changed and changed back. A new formula is in it when its
result is not empty. After a cancelled `Recalculate`, the next one reports what both calculated.
The first `Recalculate` after `XlsxReader.Load` lists the cells where SharpCell's result differs
from the one Excel saved in the file.

## Precedents and dependents

```csharp snippet=precedents
IReadOnlyList<CellRange> inputs = sheet["F2"].Precedents;   // Budget!F1, Budget!B6
IReadOnlyList<Cell> readers = sheet["B2"].Dependents;       // Budget!C2, Budget!F3, Budget!B6
```

`Cell.Precedents` lists the cells and ranges the cell's formula read in its last calculation:
references in the formula and in the names it uses, and references built while calculating, by
`INDIRECT` or `OFFSET`. A branch of `IF` that was not taken is not included, so the list is what
the value actually depends on now. A cell filled by a spill has its anchor as precedent. The
formula must be calculated: an out-of-date one throws `InvalidOperationException`.

`Cell.Dependents` lists the formula cells that read the cell, alone or within a range, and for a
spilling formula also the cells it spilled into. It throws `InvalidOperationException` while any
formula of the workbook is out of date. Both list direct links only; follow them for the rest:

```csharp snippet=affected
// Every formula an input reaches, directly or through other formulas.
static List<Cell> Affected(Cell input)
{
    var seen = new HashSet<string>();
    var found = new List<Cell>();
    var queue = new Queue<Cell>(input.Dependents);
    while (queue.Count > 0)
    {
        Cell cell = queue.Dequeue();
        if (!seen.Add(cell.ToString()))
            continue;
        found.Add(cell);
        foreach (Cell next in cell.Dependents)
            queue.Enqueue(next);
    }

    return found;
}
```

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
| `UnsupportedFormula` | A formula uses something SharpCell cannot evaluate, such as a link to another workbook. | `#NAME?` |
| `LimitExceeded` | A result was larger than a limit allows, such as the cells all spills of a workbook may cover together. | An error |
| `CustomFunctionFailure` | A [custom function](custom-functions.md) threw an exception. | `#VALUE!` |

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

## Templates in a service

Loading a file means unzipping, parsing XML and formulas, and calculating every formula.
`Workbook.Clone` copies a workbook as it is, calculated values and all, without any of that. A
service can load a template once and clone it for each request:

```csharp snippet=clone-template
// Once, at startup: load and calculate the template.
Workbook template = XlsxReader.Load(SampleTests.BudgetPath);
template.Recalculate();

// Per request, on any thread: a copy costs no file reading and no full calculation.
Parallel.ForEach(amounts, amount =>
{
    Workbook workbook = template.Clone();
    workbook["Budget"]["B2"].Value = amount;
    workbook.Recalculate();   // only what depends on B2
    results.Add(workbook["Budget"]["F2"].Value.AsNumber());
});
```

The copy is independent: changing it does not change the template, and the other way round. It
keeps the template's custom functions, and `XlsxWriter.Save` writes it into the template's file.
Parsed formulas and the file's bytes are shared, not copied.

Cloning only reads the template, so any number of threads may clone it at the same time, as long
as nothing changes or calculates the template meanwhile. Each copy, like any workbook, is for one
thread at a time.

## Culture and dates

```csharp snippet=culture
var german = new Workbook { Culture = CultureInfo.GetCultureInfo("de-DE") };
CellValue number = german.Evaluate("=VALUE(\"1,5\")");   // 1.5
```

`Workbook.Culture` controls text-to-number conversion, number-to-text conversion and text ordering. It defaults to the invariant culture, which behaves like English Excel. The process culture is never read.

`Workbook.DateSystem` says whether serial dates count from 1900 or from 1904; the .xlsx reader sets it from the file. In the 1900 system, serial number 60 is the non-existent 29 February 1900, as in Excel.
