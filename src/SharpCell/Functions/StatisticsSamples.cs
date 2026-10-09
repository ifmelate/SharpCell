using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// Reading the numbers of statistical functions. Two rule sets exist, chosen by the function name:
/// the plain functions (VAR, MEDIAN) read like SUM, the A-functions (VARA, MAXA) also count logical
/// values and text inside ranges.
/// </summary>
internal static class Samples
{
    private const int CancellationCheckInterval = 4096;

    /// <summary>
    /// Adds the numbers of arguments <paramref name="first"/> to <paramref name="end"/> (exclusive).
    /// Plain rules: ranges and arrays give numbers only; direct arguments are coerced ("2", TRUE),
    /// and text that is no number is <c>#VALUE!</c>. A-rules: ranges and arrays also give TRUE as 1
    /// and FALSE, text and text that looks like a number as 0; empty cells are skipped either way.
    /// The first error wins. As in <see cref="Aggregation.ForEach"/>, once a dirty cell has been met,
    /// the remaining referenced cells are still read so that all dirty inputs are found in one pass.
    /// </summary>
    /// <returns>Null, or the error to return.</returns>
    public static CellValue? Collect(FunctionCall call, int first, int end, bool countAll, List<double> into)
    {
        var context = call.Context;
        CellValue? error = null;
        var visited = 0;
        for (var i = first; i < end; i++)
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
                        if (error is null)
                            error = AddInRange(value, countAll, into);
                        else if (!call.MetPendingInput)
                            return error;
                    }
                }
            }
            else if (error is not null)
            {
                continue;
            }
            else if (argument.Value.Kind == CellValueKind.Array)
            {
                foreach (var element in argument.Value.AsArray())
                {
                    error = AddInRange(element, countAll, into);
                    if (error is not null)
                        break;
                }
            }
            else
            {
                // An omitted argument counts as 0, as in SUM.
                var value = argument.Value.Kind == CellValueKind.Missing ? CellValue.Number(0) : argument.Value;
                error = AddDirect(value, context, into);
            }

            if (error is not null && !call.MetPendingInput)
                return error;
        }

        return error;
    }

    /// <summary>All arguments by the plain rules into a new list.</summary>
    public static CellValue? Collect(FunctionCall call, out List<double> numbers, bool countAll = false)
    {
        numbers = [];
        return Collect(call, 0, call.Count, countAll, numbers);
    }

    private static CellValue? AddInRange(CellValue value, bool countAll, List<double> into)
    {
        switch (value.Kind)
        {
            case CellValueKind.Number:
                into.Add(value.AsNumber());
                return null;
            case CellValueKind.Error:
                return value;
            case CellValueKind.Boolean when countAll:
                into.Add(value.AsBoolean() ? 1 : 0);
                return null;
            case CellValueKind.Text when countAll:
                into.Add(0);
                return null;
            default:
                return null;
        }
    }

    private static CellValue? AddDirect(CellValue value, EvaluationContext context, List<double> into)
    {
        if (value.Kind == CellValueKind.Empty)
            return null;
        var number = Coercion.ToNumber(value, context.Culture, context.DateSystem);
        if (number.IsError)
            return number;
        into.Add(number.AsNumber());
        return null;
    }

    public static double Sum(List<double> numbers)
    {
        var total = 0.0;
        foreach (var x in numbers)
            total += x;
        return total;
    }

    public static double Mean(List<double> numbers) => Sum(numbers) / numbers.Count;

    /// <summary>The sum of squared deviations from the mean, in two passes for accuracy.</summary>
    public static double SquaredDeviations(List<double> numbers)
    {
        var mean = Mean(numbers);
        var total = 0.0;
        foreach (var x in numbers)
            total += (x - mean) * (x - mean);
        return total;
    }

    /// <summary>A sorted copy; sorting may take long, so cancellation is checked first.</summary>
    public static double[] Sorted(List<double> numbers, EvaluationContext context)
    {
        context.CancellationToken.ThrowIfCancellationRequested();
        var sorted = numbers.ToArray();
        Array.Sort(sorted);
        return sorted;
    }
}

/// <summary>
/// A range or array argument as a sequence of values in row-major order, for functions that pair
/// two arguments element by element (CORREL, SLOPE). A reference holds only its stored cells, so
/// whole columns cost what they contain.
/// </summary>
internal sealed class Sequence
{
    private Sequence(long length, int rows, int columns, List<(long Index, CellValue Value)> items)
    {
        Length = length;
        Rows = rows;
        Columns = columns;
        Items = items;
    }

    /// <summary>The number of positions, empty cells included.</summary>
    public long Length { get; }

    public int Rows { get; }

    public int Columns { get; }

    /// <summary>Positions that may hold a value, ascending; for arrays every position.</summary>
    public List<(long Index, CellValue Value)> Items { get; }

    /// <summary>A single-area reference, an array or a scalar; anything else is <c>#VALUE!</c>.</summary>
    public static bool TryRead(Operand operand, EvaluationContext context, out Sequence sequence, out CellValue error)
    {
        sequence = null!;
        error = default;
        var items = new List<(long, CellValue)>();
        if (operand.Reference is { } reference)
        {
            if (!reference.IsSingleArea)
            {
                error = CellValue.Error(ErrorKind.Value);
                return false;
            }

            var (sheet, area) = reference.Areas[0];
            var visited = 0;
            foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
            {
                if (++visited % 4096 == 0)
                    context.CancellationToken.ThrowIfCancellationRequested();
                var index = (long)(cell.Row - area.FirstRow) * area.Columns + (cell.Column - area.FirstColumn);
                items.Add((index, context.ReadCell(sheet, cell)));
            }

            sequence = new Sequence(area.CellCount, area.Rows, area.Columns, items);
            return true;
        }

        var value = operand.Value;
        if (value.Kind == CellValueKind.Array)
        {
            var array = value.AsArray();
            var columns = array.GetLength(1);
            for (var r = 0; r < array.GetLength(0); r++)
            {
                for (var c = 0; c < columns; c++)
                    items.Add(((long)r * columns + c, array[r, c]));
            }

            sequence = new Sequence(array.Length, array.GetLength(0), columns, items);
            return true;
        }

        items.Add((0, value.Kind == CellValueKind.Missing ? CellValue.Empty : value));
        sequence = new Sequence(1, 1, 1, items);
        return true;
    }

    /// <summary>
    /// The pairs of positions where both sequences hold numbers. Sequences of different lengths are
    /// <c>#N/A</c>; an error at any position of either is the result.
    /// </summary>
    /// <returns>Null, or the error to return.</returns>
    public static CellValue? Pairs(Sequence x, Sequence y, List<double> xs, List<double> ys)
    {
        if (x.Length != y.Length)
            return CellValue.Error(ErrorKind.NA);

        var i = 0;
        var j = 0;
        while (i < x.Items.Count || j < y.Items.Count)
        {
            var xi = i < x.Items.Count ? x.Items[i].Index : long.MaxValue;
            var yj = j < y.Items.Count ? y.Items[j].Index : long.MaxValue;
            var xValue = xi <= yj ? x.Items[i].Value : CellValue.Empty;
            var yValue = yj <= xi ? y.Items[j].Value : CellValue.Empty;
            if (xValue.IsError)
                return xValue;
            if (yValue.IsError)
                return yValue;
            if (xValue.Kind == CellValueKind.Number && yValue.Kind == CellValueKind.Number)
            {
                xs.Add(xValue.AsNumber());
                ys.Add(yValue.AsNumber());
            }

            if (xi <= yj)
                i++;
            if (yj <= xi)
                j++;
        }

        return null;
    }
}
