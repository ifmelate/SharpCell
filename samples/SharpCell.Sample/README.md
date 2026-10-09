# SharpCell sample

A console app that reads `budget.xlsx`, recalculates it, changes the rent, recalculates again
and evaluates a few formulas that live in no cell. The code is in `BudgetDemo.cs`; the
[getting started guide](../../docs/getting-started.md) walks through it.

```bash
dotnet run --project samples/SharpCell.Sample -f net10.0
```

Expected output:

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

`budget.xlsx` is written by `make-budget.py`. Its formulas carry no saved results, which is
why the sample recalculates before reading. A workbook saved by Excel also works: SharpCell
reads Excel's saved results and replaces them on `Recalculate`.

In your own project, reference the package instead of the source projects:

```bash
dotnet add package SharpCell.Xlsx
```
