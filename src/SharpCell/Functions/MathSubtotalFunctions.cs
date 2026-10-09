using System;
using System.Collections.Generic;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Functions;

/// <summary>
/// SUBTOTAL and AGGREGATE. Both skip cells whose formula is itself a SUBTOTAL or AGGREGATE, so
/// totals of subtotals do not count twice. SharpCell has no hidden rows: the codes and options
/// that ignore hidden rows behave like the ones that do not.
/// </summary>
internal static class MathSubtotalFunctions
{
    private const int CancellationCheckInterval = 4096;

    private const string HiddenRows = "Rows that Excel hides or filters out are included, because SharpCell does not model hidden rows.";

    // AGGREGATE's functions; SUBTOTAL uses the first eleven.
    private const int Average = 1;
    private const int Count = 2;
    private const int CountA = 3;
    private const int Max = 4;
    private const int Min = 5;
    private const int Product = 6;
    private const int StDevS = 7;
    private const int StDevP = 8;
    private const int Sum = 9;
    private const int VarS = 10;
    private const int VarP = 11;
    private const int Median = 12;
    private const int Mode = 13;
    private const int Large = 14;
    private const int Small = 15;
    private const int PercentileInc = 16;
    private const int QuartileInc = 17;
    private const int PercentileExc = 18;
    private const int QuartileExc = 19;

    public static void Register(FunctionRegistry registry)
    {
        var max = FunctionRegistry.MaxArguments;
        registry.Add(new FunctionInfo("SUBTOTAL", 2, max, [ArgumentKind.Value, ArgumentKind.Any], Subtotal)
        {
            IsVolatile = true,
            Status = FunctionStatus.KnownDeviation,
            Deviation = HiddenRows,
        });
        // The array of the array form is calculated as an array even in formulas from before
        // dynamic arrays: AGGREGATE(14,6,A1:A9/(B1:B9="x"),1) never needed Ctrl+Shift+Enter.
        registry.Add(new FunctionInfo("AGGREGATE", 3, max, [ArgumentKind.Value, ArgumentKind.Value, ArgumentKind.ArrayContext, ArgumentKind.Any], Aggregate)
        {
            IsVolatile = true,
            Status = FunctionStatus.KnownDeviation,
            Deviation = HiddenRows,
        });
    }

    // SUBTOTAL(1-11 or 101-111, ref, ...): references only; errors in them are the result.
    private static Operand Subtotal(FunctionCall call)
    {
        var code = call.Integer(0);
        if (code.IsError)
            return code;
        var function = code.AsNumber() > 100 ? code.AsNumber() - 100 : code.AsNumber();
        if (function < Average || function > VarP)
            return CellValue.Error(ErrorKind.Value);

        return Calculate(call, (int)function, first: 1, ignoreErrors: false, ignoreNested: true, k: 0);
    }

    // AGGREGATE(function, options, ref, ...) for functions 1-13; AGGREGATE(function, options,
    // array, k) for 14-19. Options 0-3 skip nested totals, options 2, 3, 6 and 7 skip errors.
    private static Operand Aggregate(FunctionCall call)
    {
        var code = call.Integer(0);
        if (code.IsError)
            return code;
        var options = call.Integer(1, 0);
        if (options.IsError)
            return options;

        var function = (int)Math.Clamp(code.AsNumber(), 0, 20);
        var option = options.AsNumber();
        if (function < Average || function > QuartileExc || option < 0 || option > 7)
            return CellValue.Error(ErrorKind.Value);

        var ignoreErrors = option is 2 or 3 or 6 or 7;
        var ignoreNested = option <= 3;
        if (function < Large)
            return Calculate(call, function, first: 2, ignoreErrors, ignoreNested, k: 0);

        if (call.Count != 4 || call.IsMissing(3))
            return CellValue.Error(ErrorKind.Value);
        var k = call.Number(3);
        if (k.IsError)
            return k;
        return Calculate(call, function, first: 2, ignoreErrors, ignoreNested, k.AsNumber(), last: 2);
    }

    private static CellValue Calculate(FunctionCall call, int function, int first, bool ignoreErrors, bool ignoreNested, double k, int last = int.MaxValue)
    {
        var numbers = new List<double>();
        long count = 0;
        CellValue? error = null;
        var failure = Visit(call, first, Math.Min(last, call.Count - 1), ignoreNested, function >= Large, value =>
        {
            switch (function)
            {
                case Count:
                    if (value.Kind == CellValueKind.Number)
                        count++;
                    return true;
                case CountA:
                    if (value.Kind != CellValueKind.Empty)
                        count++;
                    return true;
            }

            if (value.IsError)
            {
                if (ignoreErrors)
                    return true;
                error = value;
                return false;
            }

            if (value.Kind == CellValueKind.Number)
                numbers.Add(value.AsNumber());
            return true;
        });

        if (failure is { } f)
            return f;
        if (error is { } e)
            return e;
        return function is Count or CountA ? CellValue.Number(count) : Statistic(function, numbers, k, call.Context);
    }

    /// <summary>
    /// Visits the values of the arguments from <paramref name="first"/> to <paramref name="last"/>.
    /// They must be references, or arrays where <paramref name="arrays"/> allows (AGGREGATE's
    /// array form); anything else is <c>#VALUE!</c>.
    /// </summary>
    private static CellValue? Visit(FunctionCall call, int first, int last, bool ignoreNested, bool arrays, Func<CellValue, bool> visit)
    {
        var context = call.Context;
        var visited = 0;
        var visiting = true;
        for (var i = first; i <= last; i++)
        {
            if (call.IsMissing(i))
                return CellValue.Error(ErrorKind.Value);
            var argument = call[i];
            if (argument.Reference is not { } reference)
            {
                var value = argument.Value;
                if (arrays && value.Kind == CellValueKind.Array)
                {
                    foreach (var element in value.AsArray())
                    {
                        if (!visit(element))
                            return null;
                    }

                    continue;
                }

                return value.IsError ? value : CellValue.Error(ErrorKind.Value);
            }

            foreach (var (sheet, area) in reference.Areas)
            {
                foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
                {
                    if (++visited % CancellationCheckInterval == 0)
                        context.CancellationToken.ThrowIfCancellationRequested();
                    if (ignoreNested && cell.Data.Formula is { } formula && IsTotal(formula))
                        continue;

                    // After a dirty cell the rest is still read, so all dirty inputs are found in one pass.
                    var value = context.ReadCell(sheet, cell);
                    if (visiting && !visit(value))
                        visiting = false;
                    if (!visiting && !call.MetPendingInput)
                        return null;
                }
            }
        }

        return null;
    }

    // Whether a formula calls SUBTOTAL or AGGREGATE anywhere.
    private static bool IsTotal(FormulaNode formula)
    {
        var stack = new Stack<FormulaNode>();
        stack.Push(formula);
        while (stack.Count > 0)
        {
            switch (stack.Pop())
            {
                case FunctionNode f:
                    if (f.Name is "SUBTOTAL" or "AGGREGATE")
                        return true;
                    foreach (var argument in f.Arguments)
                        stack.Push(argument);
                    break;
                case CallNode c:
                    stack.Push(c.Callee);
                    foreach (var argument in c.Arguments)
                        stack.Push(argument);
                    break;
                case BinaryNode b:
                    stack.Push(b.Left);
                    stack.Push(b.Right);
                    break;
                case UnaryNode u:
                    stack.Push(u.Operand);
                    break;
                case ParenthesesNode p:
                    stack.Push(p.Inner);
                    break;
                case SpillNode s:
                    stack.Push(s.Operand);
                    break;
                case ImplicitIntersectionNode n:
                    stack.Push(n.Operand);
                    break;
            }
        }

        return false;
    }

    private static CellValue Statistic(int function, List<double> x, double k, EvaluationContext context)
    {
        var n = x.Count;
        switch (function)
        {
            case Average:
                return n == 0 ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(Total(x) / n);
            case Max:
            case Min:
                if (n == 0)
                    return CellValue.Number(0);
                var extreme = x[0];
                foreach (var value in x)
                    extreme = function == Max ? Math.Max(extreme, value) : Math.Min(extreme, value);
                return CellValue.Number(extreme);
            case Product:
                if (n == 0)
                    return CellValue.Number(0);
                var product = 1.0;
                foreach (var value in x)
                    product *= value;
                return CellValue.Number(product);
            case Sum:
                return CellValue.Number(Total(x));
            case StDevS:
            case VarS:
                if (n < 2)
                    return CellValue.Error(ErrorKind.Div0);
                var sample = SquaredDeviations(x) / (n - 1);
                return CellValue.Number(function == StDevS ? Math.Sqrt(sample) : sample);
            case StDevP:
            case VarP:
                if (n == 0)
                    return CellValue.Error(ErrorKind.Div0);
                var population = SquaredDeviations(x) / n;
                return CellValue.Number(function == StDevP ? Math.Sqrt(population) : population);
            case Median:
                if (n == 0)
                    return CellValue.Error(ErrorKind.Num);
                return CellValue.Number(PercentileInclusive(Sorted(x, context), 0.5));
            case Mode:
                return MostFrequent(x);
            case Large:
            case Small:
                var rank = Math.Ceiling(k);
                if (rank < 1 || rank > n)
                    return CellValue.Error(ErrorKind.Num);
                var sorted = Sorted(x, context);
                return CellValue.Number(function == Small ? sorted[(int)rank - 1] : sorted[n - (int)rank]);
            case PercentileInc:
            case QuartileInc:
                var p = function == QuartileInc ? Math.Truncate(k) / 4 : k;
                if (n == 0 || p < 0 || p > 1)
                    return CellValue.Error(ErrorKind.Num);
                return CellValue.Number(PercentileInclusive(Sorted(x, context), p));
            default:
                var q = function == QuartileExc ? Math.Truncate(k) / 4 : k;
                var position = q * (n + 1);
                if (q <= 0 || q >= 1 || position < 1 || position > n)
                    return CellValue.Error(ErrorKind.Num);
                var ordered = Sorted(x, context);
                var below = (int)Math.Floor(position);
                var fraction = position - below;
                return CellValue.Number(below == n
                    ? ordered[n - 1]
                    : ordered[below - 1] + fraction * (ordered[below] - ordered[below - 1]));
        }
    }

    private static double Total(List<double> x)
    {
        var total = 0.0;
        foreach (var value in x)
            total += value;
        return total;
    }

    private static double SquaredDeviations(List<double> x)
    {
        var mean = Total(x) / x.Count;
        var total = 0.0;
        foreach (var value in x)
            total += (value - mean) * (value - mean);
        return total;
    }

    private static List<double> Sorted(List<double> x, EvaluationContext context)
    {
        context.CancellationToken.ThrowIfCancellationRequested();
        var sorted = new List<double>(x);
        sorted.Sort();
        return sorted;
    }

    // Between the two nearest ranks: position p(n-1) counted from 0.
    private static double PercentileInclusive(List<double> sorted, double p)
    {
        var position = p * (sorted.Count - 1);
        var below = (int)Math.Floor(position);
        if (below >= sorted.Count - 1)
            return sorted[^1];
        return sorted[below] + (position - below) * (sorted[below + 1] - sorted[below]);
    }

    // The value that occurs most often; on a tie the one met first. No repeated value is #N/A.
    private static CellValue MostFrequent(List<double> x)
    {
        var counts = new Dictionary<double, int>();
        foreach (var value in x)
            counts[value] = counts.GetValueOrDefault(value) + 1;

        double? best = null;
        var bestCount = 1;
        foreach (var value in x)
        {
            if (counts[value] > bestCount)
            {
                best = value;
                bestCount = counts[value];
            }
        }

        return best is { } mode ? CellValue.Number(mode) : CellValue.Error(ErrorKind.NA);
    }
}
