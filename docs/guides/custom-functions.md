# Custom functions

Functions written in C# that formulas call by name, in place of VBA functions or add-ins that a
workbook expects.

## Adding a function

```csharp snippet=custom-function
workbook.Functions.Add("VAT", args =>
{
    CellValue amount = args.Number(0);   // converted as Excel converts: "100" is 100
    if (amount.IsError)
        return amount;                  // pass #VALUE! and other errors on
    return amount.AsNumber() * 0.2;
}, new FunctionOptions { MinArguments = 1, MaxArguments = 1, ScalarParameters = [0] });

sheet["A1"].Value = 100;
sheet["A2"].Value = 250;
sheet["B1"].Formula = "=VAT(A1)";      // 20
sheet["C1"].Formula = "=VAT(A1:A2)";   // a scalar parameter given a range: spills 20 and 50
workbook.Recalculate();
```

`Workbook.Functions.Add` takes a name, a body and options. Formulas call the function ignoring
case. Formulas that were `#NAME?` because the function was missing are calculated again by the
next `Recalculate`; `Remove` turns its calls back into `#NAME?`. Each workbook has its own
functions, and `Workbook.Clone` copies them.

The name has letters, digits, `.` and `_`, and starts with a letter or `_`. The name of a function
SharpCell implements is refused: their results are checked against Excel. A name Excel knows but
SharpCell does not implement, such as `WEBSERVICE`, can be added. In files, Excel writes calls of
add-in functions as `_xll.NAME(...)` and calls of functions it does not know as
`_xludf.NAME(...)`; both find the function `NAME`.

## Arguments

| Option | Meaning |
|---|---|
| `MinArguments`, `MaxArguments` | How many arguments a call may have, 0 to 255. A call outside the range is `#VALUE!`, as with Excel's functions. |
| `ScalarParameters` | Positions (from 0) of parameters that take one value. Given a range or an array, the function runs once per element and the results spill, as `LEN(A1:A10)` does. |
| `IsVolatile` | The function runs on every `Recalculate`, like `NOW`, not only when its arguments change. |

`FunctionArguments` gives the body the call's arguments:

- `args[i]` is the value: a single cell's value, a range as a `CellValueKind.Array`, an error as
  an error value (the function decides what to do with it), and `CellValue.Missing` for an
  argument left out, as the middle one of `F(1,,2)`.
- `Number(i)`, `Text(i)` and `Boolean(i)` convert the argument as Excel does and return either
  the converted value or the error to return, such as `#VALUE!` for `"abc"` read as a number.
- `Caller` is the cell whose formula makes the call; null for `Workbook.Evaluate`.
- `DateSystem` and `Culture` are the workbook's, for dates and text.

A parameter that is not scalar receives a range as an array of its values:

```csharp snippet=custom-function-range
workbook.Functions.Add("SUMSQUARES", args =>
{
    double total = 0;
    foreach (CellValue value in args[0].Kind == CellValueKind.Array ? args[0].AsArray() : new[,] { { args[0] } })
    {
        if (value.Kind == CellValueKind.Number)
            total += value.AsNumber() * value.AsNumber();
    }

    return total;
}, new FunctionOptions { MinArguments = 1, MaxArguments = 1 });

sheet["B1"].Formula = "=SUMSQUARES(A1:A3)";   // the range arrives as an array: 25
```

## Results and failures

The body returns a `CellValue`. An array result spills like any dynamic array. `Missing` and
lambda results are `#VALUE!`.

An exception thrown by the body makes the cell `#VALUE!` and adds a diagnostic of kind
`DiagnosticKind.CustomFunctionFailure`, naming the function and the exception. The rest of the
workbook calculates as usual. An `OperationCanceledException` cancels the calculation only when
its token was cancelled; otherwise it is a failure like any other exception.

## Rules for the body

- **Read cells only through the arguments.** SharpCell records which cells a formula reads, to
  know what to calculate again. A cell the body reads on its own, such as
  `args.Caller!.Worksheet["Z1"].Value`, is not recorded: the formula is not updated when it
  changes. Pass it as an argument, or make the function volatile.
- **Do not change the workbook.** Setting cells, defining names, adding sheets or functions, and
  calling `Recalculate` or `Evaluate` from inside a body throw `InvalidOperationException`, which
  turns into `#VALUE!` and a diagnostic.
- **Keep the arguments to the call.** `FunctionArguments` is valid only while the body runs.
  An array it gives is a copy, which the body may change.
- **Expect extra calls.** When a formula's inputs are out of date, SharpCell may call the body
  before calculating them, with stand-in values, and throw that result away. A body without side
  effects does not notice; one that calls a service should cache or tolerate it.
- **Stop when cancelled.** `args.CancellationToken` is the token given to `Recalculate`; a slow
  body can throw `OperationCanceledException` through it to cancel the calculation.
- **Clones share the body.** `Workbook.Clone` copies the delegate, not what it captures: a body
  used by clones on several threads must be thread-safe.

A workbook whose formulas use custom functions saves with `XlsxWriter.Save` like any other: the
results are calculated by SharpCell, so `KeepUncalculated` is not needed.
