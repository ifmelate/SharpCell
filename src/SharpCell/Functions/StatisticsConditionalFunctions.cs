using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS, on the machinery of SUMIF and SUMIFS.</summary>
internal static class StatisticsConditionalFunctions
{
    // The value range first, then (range, criterion) pairs; criteria are scalars, so an array of
    // criteria gives an array of results.
    private static readonly ArgumentKind[] TargetAndPairs = BuildTargetAndPairs();

    public static void Register(FunctionRegistry registry)
    {
        var max = FunctionRegistry.MaxArguments;
        registry.Add(new FunctionInfo("AVERAGEIF", 2, 3, [ArgumentKind.Any, ArgumentKind.Value, ArgumentKind.Any], AverageIf));
        registry.Add(new FunctionInfo("AVERAGEIFS", 3, max, TargetAndPairs, call => Ifs(call, Aggregate.Average)));
        registry.Add(new FunctionInfo("MAXIFS", 3, max, TargetAndPairs, call => Ifs(call, Aggregate.Max)));
        registry.Add(new FunctionInfo("MINIFS", 3, max, TargetAndPairs, call => Ifs(call, Aggregate.Min)));
    }

    private enum Aggregate
    {
        Average,
        Max,
        Min,
    }

    private static ArgumentKind[] BuildTargetAndPairs()
    {
        var kinds = new ArgumentKind[FunctionRegistry.MaxArguments];
        for (var i = 0; i < kinds.Length; i++)
            kinds[i] = i == 0 || (i - 1) % 2 == 0 ? ArgumentKind.Any : ArgumentKind.Value;
        return kinds;
    }

    // The average range is read from its top-left cell in the shape of the criteria range, as SUMIF does.
    private static Operand AverageIf(FunctionCall call)
    {
        if (!ValueGrid.TryCreate(call[0], out var range, out var error))
            return error;
        var criterion = Criterion.Parse(call.Value(1), call.Context.Culture, call.Context.DateSystem);
        var target = range;
        if (call.Has(2))
        {
            if (!ValueGrid.TryCreate(call[2], out var values, out error))
                return error;
            if (values.Resized(range.Rows, range.Columns, call.Context) is not { } resized)
                return CellValue.Error(ErrorKind.Value);
            target = resized;
        }

        return Apply(call, [(range, criterion)], target, Aggregate.Average);
    }

    private static Operand Ifs(FunctionCall call, Aggregate aggregate)
    {
        if (!ValueGrid.TryCreate(call[0], out var target, out var error))
            return error;
        if (!Conditional.TryReadPairs(call, 1, out var conditions, out error))
            return error;
        return Apply(call, conditions, target, aggregate);
    }

    // Only numbers in matching cells count; an error in a matching cell is the result. No number
    // at all is #DIV/0! for the average and 0 for MAXIFS and MINIFS.
    private static Operand Apply(FunctionCall call, List<(ValueGrid, Criterion)> conditions, ValueGrid target, Aggregate aggregate)
    {
        var total = 0.0;
        var count = 0L;
        var extreme = 0.0;
        CellValue? firstError = null;
        var failure = Conditional.Visit(call, conditions, target, value =>
        {
            if (value.IsError)
            {
                firstError ??= value;
                return;
            }

            if (value.Kind != CellValueKind.Number)
                return;
            var x = value.AsNumber();
            extreme = count == 0 ? x : aggregate == Aggregate.Max ? Math.Max(extreme, x) : Math.Min(extreme, x);
            total += x;
            count++;
        }, _ => { });

        if ((failure ?? firstError) is { } result)
            return result;
        return aggregate switch
        {
            Aggregate.Average => count == 0 ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(total / count),
            _ => CellValue.Number(count == 0 ? 0 : extreme),
        };
    }
}
