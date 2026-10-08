using System;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

internal enum ValueSource
{
    /// <summary>Typed directly as an argument: <c>SUM(1,"2",TRUE)</c>.</summary>
    Direct,

    /// <summary>A cell of a referenced range. Only existing cells are visited.</summary>
    Reference,

    /// <summary>An element of an array argument.</summary>
    Array,
}

/// <summary>Walks the values of SUM-like arguments. Many Excel functions treat the three sources differently.</summary>
internal static class Aggregation
{
    private const int CancellationCheckInterval = 4096;

    /// <summary>
    /// Visits every value; the visitor returns false to stop early. Once a dirty cell has been met,
    /// the remaining referenced cells are still read (not visited) so that all dirty inputs are
    /// collected in one pass: a total above 100 000 dirty formulas is then evaluated twice, not
    /// 100 000 times.
    /// </summary>
    public static void ForEach(FunctionCall call, Func<CellValue, ValueSource, bool> visit)
    {
        var context = call.Context;
        var visited = 0;
        var visiting = true;
        for (var i = 0; i < call.Count; i++)
        {
            var argument = call[i];
            if (argument.Reference is { } reference)
            {
                foreach (var (sheet, area) in reference.Areas)
                {
                    foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
                    {
                        if (++visited % CancellationCheckInterval == 0)
                            context.CancellationToken.ThrowIfCancellationRequested();
                        var value = context.ReadCell(sheet, cell);
                        if (visiting && !visit(value, ValueSource.Reference))
                            visiting = false;
                        if (!visiting && !call.MetPendingInput)
                            return;
                    }
                }
            }
            else if (!visiting)
            {
                continue;
            }
            else if (argument.Value.Kind == CellValueKind.Array)
            {
                foreach (var element in argument.Value.AsArray())
                {
                    if (!visit(element, ValueSource.Array))
                    {
                        visiting = false;
                        break;
                    }
                }
            }
            else
            {
                // An omitted argument counts as 0: SUM(1,) is 1, COUNT(1,) is 2.
                var value = argument.Value.Kind == CellValueKind.Missing ? CellValue.Number(0) : argument.Value;
                visiting = visit(value, ValueSource.Direct);
            }

            if (!visiting && !call.MetPendingInput)
                return;
        }
    }

    /// <summary>
    /// Numbers for SUM, AVERAGE, MIN and MAX. Ranges and arrays contribute numbers only; direct
    /// arguments are coerced (text "2", TRUE). The first error wins.
    /// </summary>
    public static CellValue Numbers(FunctionCall call, Action<double> add)
    {
        CellValue? error = null;
        ForEach(call, (value, source) =>
        {
            if (value.IsError)
            {
                error = value;
                return false;
            }

            if (value.Kind == CellValueKind.Number)
            {
                add(value.AsNumber());
                return true;
            }

            if (source != ValueSource.Direct || value.Kind == CellValueKind.Empty)
                return true;

            var number = Coercion.ToNumber(value, call.Context.Culture);
            if (number.IsError)
            {
                error = number;
                return false;
            }

            add(number.AsNumber());
            return true;
        });
        return error ?? CellValue.Empty;
    }
}
