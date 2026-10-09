using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// Descriptive statistics over SUM-like arguments: averages, variances, moments, the median and
/// the mode, and COUNTBLANK. The A-variants (AVERAGEA, VARA...) count TRUE as 1 and text as 0.
/// </summary>
internal static class StatisticsFunctions
{
    private static readonly ArgumentKind[] AnyArguments = [ArgumentKind.Any];

    public static void Register(FunctionRegistry registry)
    {
        var max = FunctionRegistry.MaxArguments;
        void Add(string name, FunctionBody body) => registry.Add(new FunctionInfo(name, 1, max, AnyArguments, body));

        Add("AVERAGEA", call => Average(call));
        Add("MAXA", call => Extreme(call, Math.Max));
        Add("MINA", call => Extreme(call, Math.Min));

        Add("VAR.S", call => Variance(call, sample: true, countAll: false, root: false));
        Add("VAR", call => Variance(call, sample: true, countAll: false, root: false));
        Add("VAR.P", call => Variance(call, sample: false, countAll: false, root: false));
        Add("VARP", call => Variance(call, sample: false, countAll: false, root: false));
        Add("VARA", call => Variance(call, sample: true, countAll: true, root: false));
        Add("VARPA", call => Variance(call, sample: false, countAll: true, root: false));
        Add("STDEV.S", call => Variance(call, sample: true, countAll: false, root: true));
        Add("STDEV", call => Variance(call, sample: true, countAll: false, root: true));
        Add("STDEV.P", call => Variance(call, sample: false, countAll: false, root: true));
        Add("STDEVP", call => Variance(call, sample: false, countAll: false, root: true));
        Add("STDEVA", call => Variance(call, sample: true, countAll: true, root: true));
        Add("STDEVPA", call => Variance(call, sample: false, countAll: true, root: true));

        Add("AVEDEV", AverageDeviation);
        Add("DEVSQ", SquaredDeviations);
        Add("GEOMEAN", GeometricMean);
        Add("HARMEAN", HarmonicMean);
        Add("KURT", Kurtosis);
        Add("SKEW", call => Skewness(call, population: false));
        Add("SKEW.P", call => Skewness(call, population: true));
        Add("MEDIAN", Median);
        Add("MODE.SNGL", call => Mode(call, multiple: false));
        Add("MODE", call => Mode(call, multiple: false));
        Add("MODE.MULT", call => Mode(call, multiple: true));

        registry.Add(new FunctionInfo("COUNTBLANK", 1, 1, AnyArguments, CountBlank));
    }

    private static Operand Average(FunctionCall call)
    {
        if (Samples.Collect(call, out var numbers, countAll: true) is { } error)
            return error;
        return numbers.Count == 0 ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(Samples.Mean(numbers));
    }

    // Like MAX and MIN, no numbers at all give 0.
    private static Operand Extreme(FunctionCall call, Func<double, double, double> pick)
    {
        if (Samples.Collect(call, out var numbers, countAll: true) is { } error)
            return error;
        if (numbers.Count == 0)
            return CellValue.Number(0);
        var result = numbers[0];
        for (var i = 1; i < numbers.Count; i++)
            result = pick(result, numbers[i]);
        return CellValue.Number(result);
    }

    // A sample needs two numbers, a population one; fewer is #DIV/0!.
    private static Operand Variance(FunctionCall call, bool sample, bool countAll, bool root)
    {
        if (Samples.Collect(call, out var numbers, countAll) is { } error)
            return error;
        var divisor = sample ? numbers.Count - 1 : numbers.Count;
        if (divisor <= 0)
            return CellValue.Error(ErrorKind.Div0);
        var variance = Samples.SquaredDeviations(numbers) / divisor;
        return CellValue.Number(root ? Math.Sqrt(variance) : variance);
    }

    private static Operand AverageDeviation(FunctionCall call)
    {
        if (Samples.Collect(call, out var numbers) is { } error)
            return error;
        if (numbers.Count == 0)
            return CellValue.Error(ErrorKind.Num);
        var mean = Samples.Mean(numbers);
        var total = 0.0;
        foreach (var x in numbers)
            total += Math.Abs(x - mean);
        return CellValue.Number(total / numbers.Count);
    }

    private static Operand SquaredDeviations(FunctionCall call)
    {
        if (Samples.Collect(call, out var numbers) is { } error)
            return error;
        return numbers.Count == 0 ? CellValue.Error(ErrorKind.Num) : CellValue.Number(Samples.SquaredDeviations(numbers));
    }

    // Only positive numbers have a geometric or harmonic mean. The geometric mean goes through
    // logarithms: a product of many numbers would overflow.
    private static Operand GeometricMean(FunctionCall call)
    {
        if (Samples.Collect(call, out var numbers) is { } error)
            return error;
        if (numbers.Count == 0)
            return CellValue.Error(ErrorKind.Num);
        var logs = 0.0;
        foreach (var x in numbers)
        {
            if (x <= 0)
                return CellValue.Error(ErrorKind.Num);
            logs += Math.Log(x);
        }

        return CellValue.Number(Math.Exp(logs / numbers.Count));
    }

    private static Operand HarmonicMean(FunctionCall call)
    {
        if (Samples.Collect(call, out var numbers) is { } error)
            return error;
        if (numbers.Count == 0)
            return CellValue.Error(ErrorKind.Num);
        var reciprocals = 0.0;
        foreach (var x in numbers)
        {
            if (x <= 0)
                return CellValue.Error(ErrorKind.Num);
            reciprocals += 1 / x;
        }

        return CellValue.Number(numbers.Count / reciprocals);
    }

    // Excel's sample excess kurtosis; fewer than four numbers or no spread is #DIV/0!.
    private static Operand Kurtosis(FunctionCall call)
    {
        if (Samples.Collect(call, out var numbers) is { } error)
            return error;
        double n = numbers.Count;
        if (n < 4)
            return CellValue.Error(ErrorKind.Div0);
        var mean = Samples.Mean(numbers);
        var deviation = Math.Sqrt(Samples.SquaredDeviations(numbers) / (n - 1));
        if (deviation == 0)
            return CellValue.Error(ErrorKind.Div0);
        var fourth = 0.0;
        foreach (var x in numbers)
        {
            var z = (x - mean) / deviation;
            fourth += z * z * z * z;
        }

        var result = n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * fourth - 3 * (n - 1) * (n - 1) / ((n - 2) * (n - 3));
        return CellValue.Number(result);
    }

    // SKEW uses the sample deviation and needs three numbers; SKEW.P the population deviation.
    private static Operand Skewness(FunctionCall call, bool population)
    {
        if (Samples.Collect(call, out var numbers) is { } error)
            return error;
        double n = numbers.Count;
        if (n < (population ? 1 : 3))
            return CellValue.Error(ErrorKind.Div0);
        var mean = Samples.Mean(numbers);
        var deviation = Math.Sqrt(Samples.SquaredDeviations(numbers) / (population ? n : n - 1));
        if (deviation == 0)
            return CellValue.Error(ErrorKind.Div0);
        var third = 0.0;
        foreach (var x in numbers)
        {
            var z = (x - mean) / deviation;
            third += z * z * z;
        }

        return CellValue.Number(population ? third / n : n / ((n - 1) * (n - 2)) * third);
    }

    private static Operand Median(FunctionCall call)
    {
        if (Samples.Collect(call, out var numbers) is { } error)
            return error;
        if (numbers.Count == 0)
            return CellValue.Error(ErrorKind.Num);
        var sorted = Samples.Sorted(numbers, call.Context);
        var middle = sorted.Length / 2;
        return CellValue.Number(sorted.Length % 2 == 1 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2);
    }

    // The most frequent number; among equally frequent ones the first in the data comes first.
    // No number occurring twice is #N/A. MODE.MULT returns all of them as a column.
    private static Operand Mode(FunctionCall call, bool multiple)
    {
        if (Samples.Collect(call, out var numbers) is { } error)
            return error;

        var counts = new Dictionary<double, int>();
        var order = new List<double>();
        var best = 0;
        foreach (var x in numbers)
        {
            counts.TryGetValue(x, out var count);
            if (count == 0)
                order.Add(x);
            counts[x] = ++count;
            best = Math.Max(best, count);
        }

        if (best < 2)
            return CellValue.Error(ErrorKind.NA);

        var modes = new List<double>();
        foreach (var x in order)
        {
            if (counts[x] != best)
                continue;
            if (!multiple)
                return CellValue.Number(x);
            modes.Add(x);
        }

        var result = new CellValue[modes.Count, 1];
        for (var i = 0; i < modes.Count; i++)
            result[i, 0] = CellValue.Number(modes[i]);
        return CellValue.Array(result);
    }

    // Empty cells and cells holding empty text (="" included). Only stored cells are read, so a
    // whole column costs what it contains. An array argument counts its empty elements.
    private static Operand CountBlank(FunctionCall call)
    {
        var argument = call[0];
        if (argument.Reference is not { } reference)
        {
            var value = argument.Value;
            if (value.Kind != CellValueKind.Array)
                return value.IsError ? value : CellValue.Number(IsBlank(value) ? 1 : 0);
            var blanks = 0;
            foreach (var element in value.AsArray())
            {
                if (IsBlank(element))
                    blanks++;
            }

            return CellValue.Number(blanks);
        }

        if (!reference.IsSingleArea)
            return CellValue.Error(ErrorKind.Value);

        var context = call.Context;
        var (sheet, area) = reference.Areas[0];
        var filled = 0L;
        var visited = 0;
        foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
        {
            if (++visited % 4096 == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            if (!IsBlank(context.ReadCell(sheet, cell)))
                filled++;
        }

        return CellValue.Number(area.CellCount - filled);
    }

    private static bool IsBlank(CellValue value) =>
        value.Kind is CellValueKind.Empty or CellValueKind.Missing || (value.Kind == CellValueKind.Text && value.AsText().Length == 0);
}
