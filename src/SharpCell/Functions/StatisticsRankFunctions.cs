using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// Order statistics: LARGE, SMALL, percentiles and quartiles, percent ranks, RANK, TRIMMEAN and
/// FREQUENCY. The data argument reads like SUM's (numbers only in ranges and arrays, a direct
/// value coerced); the position argument is a scalar, so an array of positions gives an array.
/// </summary>
internal static class StatisticsRankFunctions
{
    // Before dynamic arrays, Excel already calculated these array parameters as arrays:
    // LARGE(A1:A9*2,1) needed no Ctrl+Shift+Enter.
    private static readonly ArgumentKind[] DataAndValues = [ArgumentKind.ArrayContext, ArgumentKind.Value];

    public static void Register(FunctionRegistry registry)
    {
        void Add(string name, int min, int max, ArgumentKind[] kinds, FunctionBody body) =>
            registry.Add(new FunctionInfo(name, min, max, kinds, body));

        Add("LARGE", 2, 2, DataAndValues, call => Kth(call, largest: true));
        Add("SMALL", 2, 2, DataAndValues, call => Kth(call, largest: false));
        Add("PERCENTILE.INC", 2, 2, DataAndValues, call => Percentile(call, exclusive: false, quartile: false));
        Add("PERCENTILE", 2, 2, DataAndValues, call => Percentile(call, exclusive: false, quartile: false));
        Add("PERCENTILE.EXC", 2, 2, DataAndValues, call => Percentile(call, exclusive: true, quartile: false));
        Add("QUARTILE.INC", 2, 2, DataAndValues, call => Percentile(call, exclusive: false, quartile: true));
        Add("QUARTILE", 2, 2, DataAndValues, call => Percentile(call, exclusive: false, quartile: true));
        Add("QUARTILE.EXC", 2, 2, DataAndValues, call => Percentile(call, exclusive: true, quartile: true));
        Add("PERCENTRANK.INC", 2, 3, DataAndValues, call => PercentRank(call, exclusive: false));
        Add("PERCENTRANK", 2, 3, DataAndValues, call => PercentRank(call, exclusive: false));
        Add("PERCENTRANK.EXC", 2, 3, DataAndValues, call => PercentRank(call, exclusive: true));
        Add("TRIMMEAN", 2, 2, DataAndValues, TrimMean);

        // RANK's list is a reference, as in Excel; the number and the order are scalars.
        ArgumentKind[] rank = [ArgumentKind.Value, ArgumentKind.Any, ArgumentKind.Value];
        Add("RANK.EQ", 2, 3, rank, call => Rank(call, average: false));
        Add("RANK", 2, 3, rank, call => Rank(call, average: false));
        Add("RANK.AVG", 2, 3, rank, call => Rank(call, average: true));

        Add("FREQUENCY", 2, 2, [ArgumentKind.ArrayContext], Frequency);
    }

    private static CellValue? Data(FunctionCall call, int index, out List<double> numbers)
    {
        numbers = [];
        return Samples.Collect(call, index, index + 1, false, numbers);
    }

    // Floor that does not fall one short when floating-point noise puts a whole number just below
    // itself: 0.58*100 is 57.99999999999999.
    private static double ApproxFloor(double x)
    {
        var nearest = Math.Round(x);
        return NumberComparer.AreEqual(x, nearest) ? nearest : Math.Floor(x);
    }

    // A fractional k is rounded up, as LibreOffice does; below 1 or beyond the count is #NUM!.
    private static Operand Kth(FunctionCall call, bool largest)
    {
        if (Data(call, 0, out var numbers) is { } error)
            return error;
        var k = call.Number(1);
        if (k.IsError)
            return k;
        var position = Math.Ceiling(k.AsNumber());
        if (position < 1 || position > numbers.Count)
            return CellValue.Error(ErrorKind.Num);
        var sorted = Samples.Sorted(numbers, call.Context);
        var index = (int)position - 1;
        return CellValue.Number(largest ? sorted[^(index + 1)] : sorted[index]);
    }

    // PERCENTILE.INC interpolates at k*(n-1) from the smallest (k in [0;1]); PERCENTILE.EXC at
    // k*(n+1)-1, which must fall inside the data. Quartiles are the percentiles of quart/4, the
    // quart truncated: 0 to 4 inclusive, 1 to 3 exclusive.
    private static Operand Percentile(FunctionCall call, bool exclusive, bool quartile)
    {
        if (Data(call, 0, out var numbers) is { } error)
            return error;
        var argument = quartile ? call.Integer(1) : call.Number(1);
        if (argument.IsError)
            return argument;

        var k = argument.AsNumber();
        if (quartile)
        {
            if (exclusive ? k <= 0 || k >= 4 : k < 0 || k > 4)
                return CellValue.Error(ErrorKind.Num);
            k /= 4;
        }

        var n = numbers.Count;
        if (n == 0)
            return CellValue.Error(ErrorKind.Num);

        double position;
        if (exclusive)
        {
            position = k * (n + 1) - 1;
            if (k <= 0 || k >= 1 || position < 0 || position > n - 1)
                return CellValue.Error(ErrorKind.Num);
        }
        else
        {
            if (k < 0 || k > 1)
                return CellValue.Error(ErrorKind.Num);
            position = k * (n - 1);
        }

        var sorted = Samples.Sorted(numbers, call.Context);
        var low = (int)Math.Floor(position);
        var fraction = position - low;
        return CellValue.Number(low + 1 < n ? sorted[low] + fraction * (sorted[low + 1] - sorted[low]) : sorted[low]);
    }

    // The share of the data below x: (count below)/(n-1) inclusive, (count below+1)/(n+1)
    // exclusive, interpolated between neighbours when x is not in the data. Outside the data is
    // #N/A. The result is truncated (not rounded) to the significance, 3 digits by default.
    private static Operand PercentRank(FunctionCall call, bool exclusive)
    {
        if (Data(call, 0, out var numbers) is { } error)
            return error;
        var x = call.Number(1);
        if (x.IsError)
            return x;
        var significance = call.Integer(2, 3);
        if (significance.IsError)
            return significance;
        var digits = significance.AsNumber();
        if (digits < 1)
            return CellValue.Error(ErrorKind.Num);

        var n = numbers.Count;
        if (n == 0)
            return CellValue.Error(ErrorKind.Num);
        var sorted = Samples.Sorted(numbers, call.Context);
        var value = x.AsNumber();
        if (value < sorted[0] || value > sorted[^1])
            return CellValue.Error(ErrorKind.NA);

        var below = LowerBound(sorted, value);
        double rank;
        if (sorted[below] == value)
        {
            rank = exclusive ? (below + 1.0) / (n + 1) : n == 1 ? 1 : (double)below / (n - 1);
        }
        else
        {
            // sorted[below - 1] < value < sorted[below].
            var fraction = (value - sorted[below - 1]) / (sorted[below] - sorted[below - 1]);
            rank = exclusive ? (below + fraction) / (n + 1) : (below - 1 + fraction) / (n - 1);
        }

        var scale = Math.Pow(10, Math.Min(digits, 15));
        return CellValue.Number(ApproxFloor(rank * scale) / scale);
    }

    // The index of the first element not less than value.
    private static int LowerBound(double[] sorted, double value)
    {
        int low = 0, high = sorted.Length;
        while (low < high)
        {
            var middle = (low + high) / 2;
            if (sorted[middle] < value)
                low = middle + 1;
            else
                high = middle;
        }

        return low;
    }

    // The mean without the given share of the data, half from each end, rounded down to an even
    // number of points.
    private static Operand TrimMean(FunctionCall call)
    {
        if (Data(call, 0, out var numbers) is { } error)
            return error;
        var percent = call.Number(1);
        if (percent.IsError)
            return percent;
        var share = percent.AsNumber();
        if (share < 0 || share >= 1 || numbers.Count == 0)
            return CellValue.Error(ErrorKind.Num);

        var sorted = Samples.Sorted(numbers, call.Context);
        var trim = (int)ApproxFloor(sorted.Length * share / 2);
        var total = 0.0;
        for (var i = trim; i < sorted.Length - trim; i++)
            total += sorted[i];
        return CellValue.Number(total / (sorted.Length - 2 * trim));
    }

    // RANK.EQ: 1 + the count of numbers before it in the order; RANK.AVG averages the positions
    // of ties. A number not in the list is #N/A.
    private static Operand Rank(FunctionCall call, bool average)
    {
        var number = call.Number(0);
        if (number.IsError)
            return number;
        var order = call.Number(2, 0);
        if (order.IsError)
            return order;
        if (Data(call, 1, out var numbers) is { } error)
            return error;

        var value = number.AsNumber();
        var ascending = order.AsNumber() != 0;
        int before = 0, equal = 0;
        foreach (var x in numbers)
        {
            if (x == value)
                equal++;
            else if (ascending ? x < value : x > value)
                before++;
        }

        if (equal == 0)
            return CellValue.Error(ErrorKind.NA);
        return CellValue.Number(average ? before + (equal + 1) / 2.0 : before + 1);
    }

    // A column with one count per bin, in the bins' own order, and a last count for the numbers
    // above every bin. A number goes to the smallest bin not below it; of equal bins, the first.
    private static Operand Frequency(FunctionCall call)
    {
        if (Data(call, 0, out var data) is { } error)
            return error;
        if (Data(call, 1, out var bins) is { } binError)
            return binError;
        if (bins.Count + 1L > Evaluator.MaxArrayCells)
            return CellValue.Error(ErrorKind.Num);

        var order = new int[bins.Count];
        for (var i = 0; i < order.Length; i++)
            order[i] = i;
        var keys = bins.ToArray();
        call.Context.CancellationToken.ThrowIfCancellationRequested();

        // A stable sort, so that of equal bins the first in the argument gets the counts.
        Array.Sort(order, (a, b) => keys[a] != keys[b] ? keys[a].CompareTo(keys[b]) : a.CompareTo(b));
        var sortedBins = new double[order.Length];
        for (var i = 0; i < order.Length; i++)
            sortedBins[i] = keys[order[i]];

        var counts = new long[bins.Count + 1];
        foreach (var x in data)
        {
            var bin = LowerBound(sortedBins, x);
            counts[bin < order.Length ? order[bin] : bins.Count]++;
        }

        var result = new CellValue[counts.Length, 1];
        for (var i = 0; i < counts.Length; i++)
            result[i, 0] = CellValue.Number(counts[i]);
        return CellValue.Array(result);
    }
}
