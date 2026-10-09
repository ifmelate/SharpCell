using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// Hypothesis tests over samples (T.TEST, F.TEST, CHISQ.TEST, Z.TEST) and PROB. The samples are
/// array parameters: in a formula from before dynamic arrays they are still calculated as arrays,
/// as Excel did. Only numbers in the samples count; text, logical values and empty cells are
/// skipped, and an error in a sample is the result.
/// </summary>
internal static class DistributionTestFunctions
{
    private const int CancellationCheckInterval = 4096;

    private static readonly ArgumentKind[] TwoSamples = [ArgumentKind.ArrayContext, ArgumentKind.ArrayContext, ArgumentKind.Value];
    private static readonly ArgumentKind[] OneSample = [ArgumentKind.ArrayContext, ArgumentKind.Value];

    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("T.TEST", 4, 4, TwoSamples, TTest));
        registry.Add(new FunctionInfo("TTEST", 4, 4, TwoSamples, TTest));
        registry.Add(new FunctionInfo("F.TEST", 2, 2, TwoSamples, FTest));
        registry.Add(new FunctionInfo("FTEST", 2, 2, TwoSamples, FTest));
        registry.Add(new FunctionInfo("CHISQ.TEST", 2, 2, TwoSamples, ChiSquareTest));
        registry.Add(new FunctionInfo("CHITEST", 2, 2, TwoSamples, ChiSquareTest));
        registry.Add(new FunctionInfo("Z.TEST", 2, 3, OneSample, ZTest));
        registry.Add(new FunctionInfo("ZTEST", 2, 3, OneSample, ZTest));
        registry.Add(new FunctionInfo("PROB", 3, 4, TwoSamples, Prob));
    }

    // -----------------------------------------------------------------------------------------
    // Reading samples

    /// <summary>The numbers of a range or array argument, in row-major order.</summary>
    private static CellValue? Numbers(FunctionCall call, int index, List<double> numbers)
    {
        if (!ValueGrid.TryCreate(call[index], out var grid, out var error))
            return error;

        var context = call.Context;
        var visited = 0;
        foreach (var (row, column) in Occupied(grid))
        {
            if (++visited % CancellationCheckInterval == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            var value = grid.Get(row, column, context);
            if (value.IsError)
                return value;
            if (value.Kind == CellValueKind.Number)
                numbers.Add(value.AsNumber());
        }

        return null;
    }

    /// <summary>
    /// Visits the values of two arguments position by position in row-major order (the n-th cell
    /// of one with the n-th cell of the other), skipping positions empty in both. Arguments with
    /// different numbers of cells are <c>#N/A</c>.
    /// </summary>
    private static CellValue? Pairs(FunctionCall call, int firstIndex, int secondIndex, Func<CellValue, CellValue, CellValue?> visit)
    {
        if (!ValueGrid.TryCreate(call[firstIndex], out var first, out var error))
            return error;
        if (!ValueGrid.TryCreate(call[secondIndex], out var second, out error))
            return error;
        if ((long)first.Rows * first.Columns != (long)second.Rows * second.Columns)
            return CellValue.Error(ErrorKind.NA);

        var linear = new HashSet<long>();
        foreach (var (row, column) in Occupied(first))
            linear.Add((long)row * first.Columns + column);
        foreach (var (row, column) in Occupied(second))
            linear.Add((long)row * second.Columns + column);
        var positions = new List<long>(linear);
        positions.Sort();

        var context = call.Context;
        var visited = 0;
        foreach (var position in positions)
        {
            if (++visited % CancellationCheckInterval == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            var a = first.Get((int)(position / first.Columns), (int)(position % first.Columns), context);
            var b = second.Get((int)(position / second.Columns), (int)(position % second.Columns), context);
            if (a.IsError)
                return a;
            if (b.IsError)
                return b;
            if (visit(a, b) is { } failure)
                return failure;
        }

        return null;
    }

    private static List<(int Row, int Column)> Occupied(ValueGrid grid)
    {
        var set = new HashSet<(int, int)>();
        grid.AddOccupied(set);
        var list = new List<(int Row, int Column)>(set);
        list.Sort();
        return list;
    }

    private static double Mean(List<double> values)
    {
        var sum = 0.0;
        foreach (var value in values)
            sum += value;
        return sum / values.Count;
    }

    // The sample variance, by two passes.
    private static double Variance(List<double> values, double mean)
    {
        var sum = 0.0;
        foreach (var value in values)
            sum += (value - mean) * (value - mean);
        return sum / (values.Count - 1);
    }

    // -----------------------------------------------------------------------------------------
    // Tests

    // T.TEST(array1, array2, tails, type): type 1 paired, 2 equal variances, 3 unequal (Welch).
    private static Operand TTest(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var tails = args.Integer(2);
        var type = args.Integer(3);
        if (args.Failed)
            return args.Error;
        if ((tails != 1 && tails != 2) || type < 1 || type > 3)
            return CellValue.Error(ErrorKind.Num);

        double t, freedom;
        if (type == 1)
        {
            // Pairs where both values are numbers.
            var differences = new List<double>();
            var failure = Pairs(call, 0, 1, (a, b) =>
            {
                if (a.Kind == CellValueKind.Number && b.Kind == CellValueKind.Number)
                    differences.Add(a.AsNumber() - b.AsNumber());
                return null;
            });
            if (failure is { } error)
                return error;
            if (differences.Count < 2)
                return CellValue.Error(ErrorKind.Div0);

            var mean = Mean(differences);
            var variance = Variance(differences, mean);
            if (variance == 0)
                return CellValue.Error(ErrorKind.Div0);
            t = mean / Math.Sqrt(variance / differences.Count);
            freedom = differences.Count - 1;
        }
        else
        {
            var first = new List<double>();
            var second = new List<double>();
            if (Numbers(call, 0, first) is { } firstError)
                return firstError;
            if (Numbers(call, 1, second) is { } secondError)
                return secondError;
            if (first.Count < 2 || second.Count < 2)
                return CellValue.Error(ErrorKind.Div0);

            double n1 = first.Count, n2 = second.Count;
            var m1 = Mean(first);
            var m2 = Mean(second);
            var v1 = Variance(first, m1);
            var v2 = Variance(second, m2);
            if (type == 2)
            {
                freedom = n1 + n2 - 2;
                var pooled = ((n1 - 1) * v1 + (n2 - 1) * v2) / freedom;
                var scale = Math.Sqrt(pooled * (1 / n1 + 1 / n2));
                if (scale == 0)
                    return CellValue.Error(ErrorKind.Div0);
                t = (m1 - m2) / scale;
            }
            else
            {
                var s1 = v1 / n1;
                var s2 = v2 / n2;
                if (s1 + s2 == 0)
                    return CellValue.Error(ErrorKind.Div0);
                t = (m1 - m2) / Math.Sqrt(s1 + s2);
                freedom = (s1 + s2) * (s1 + s2) / (s1 * s1 / (n1 - 1) + s2 * s2 / (n2 - 1));
            }
        }

        return CellValue.Number(tails * DistributionFunctions.StudentTail(t, freedom));
    }

    // F.TEST(array1, array2): the two-tailed probability that the variances do not differ.
    private static Operand FTest(FunctionCall call)
    {
        var first = new List<double>();
        var second = new List<double>();
        if (Numbers(call, 0, first) is { } firstError)
            return firstError;
        if (Numbers(call, 1, second) is { } secondError)
            return secondError;
        if (first.Count < 2 || second.Count < 2)
            return CellValue.Error(ErrorKind.Div0);

        var v1 = Variance(first, Mean(first));
        var v2 = Variance(second, Mean(second));
        if (v1 == 0 || v2 == 0)
            return CellValue.Error(ErrorKind.Div0);

        var f = v1 / v2;
        double d1 = first.Count - 1, d2 = second.Count - 1;
        var lower = DistributionFunctions.FisherCdf(f, d1, d2, upper: false);
        var upper = DistributionFunctions.FisherCdf(f, d1, d2, upper: true);
        return CellValue.Number(2 * Math.Min(lower, upper));
    }

    // CHISQ.TEST(actual, expected): the right tail of sum (A - E)^2 / E, with (r - 1)(c - 1)
    // degrees of freedom for a table and n - 1 for a single row or column.
    private static Operand ChiSquareTest(FunctionCall call)
    {
        var statistic = 0.0;
        var failure = Pairs(call, 0, 1, (actual, expected) =>
        {
            if (actual.Kind != CellValueKind.Number || expected.Kind != CellValueKind.Number)
                return null;
            var e = expected.AsNumber();
            if (e < 0)
                return CellValue.Error(ErrorKind.Num);
            if (e == 0)
                return CellValue.Error(ErrorKind.Div0);
            var difference = actual.AsNumber() - e;
            statistic += difference * difference / e;
            return null;
        });
        if (failure is { } error)
            return error;

        ValueGrid.TryCreate(call[0], out var shape, out _);
        double rows = shape.Rows, columns = shape.Columns;
        var freedom = rows > 1 && columns > 1 ? (rows - 1) * (columns - 1) : rows * columns - 1;
        if (freedom < 1)
            return CellValue.Error(ErrorKind.NA);
        return CellValue.Number(Special.GammaQ(freedom / 2, statistic / 2));
    }

    // Z.TEST(array, x, [sigma]): P(Z > (mean - x) / (sigma / sqrt n)); without sigma, the sample
    // standard deviation.
    private static Operand ZTest(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(1);
        var sigma = call.Has(2) ? args.Number(2) : double.NaN;
        if (args.Failed)
            return args.Error;

        var values = new List<double>();
        if (Numbers(call, 0, values) is { } error)
            return error;
        if (values.Count == 0)
            return CellValue.Error(ErrorKind.NA);

        var mean = Mean(values);
        if (call.Has(2))
        {
            if (sigma <= 0)
                return CellValue.Error(ErrorKind.Num);
        }
        else
        {
            if (values.Count < 2)
                return CellValue.Error(ErrorKind.Div0);
            sigma = Math.Sqrt(Variance(values, mean));
            if (sigma == 0)
                return CellValue.Error(ErrorKind.Div0);
        }

        var z = (mean - x) / (sigma / Math.Sqrt(values.Count));
        return CellValue.Number(Special.NormalCdf(-z));
    }

    // PROB(x_range, prob_range, lower, [upper]): the probability that a value lies in
    // [lower, upper]. The probabilities must lie in [0, 1] and add up to 1.
    private static Operand Prob(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var lower = args.Number(2);
        var upper = call.Has(3) ? args.Number(3) : lower;
        if (args.Failed)
            return args.Error;

        var total = 0.0;
        var inRange = 0.0;
        var any = false;
        var failure = Pairs(call, 0, 1, (value, probability) =>
        {
            if (probability.Kind != CellValueKind.Number)
                return null;
            var p = probability.AsNumber();
            if (p < 0 || p > 1)
                return CellValue.Error(ErrorKind.Num);
            any = true;
            total += p;
            if (value.Kind == CellValueKind.Number && value.AsNumber() >= lower && value.AsNumber() <= upper)
                inRange += p;
            return null;
        });
        if (failure is { } error)
            return error;
        if (!any)
            return CellValue.Error(ErrorKind.Div0);
        if (!NumberComparer.AreEqual(total, 1))
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(inRange);
    }
}
