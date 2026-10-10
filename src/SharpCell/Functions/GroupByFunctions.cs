using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>PERCENTOF; GROUPBY and PIVOTBY build on it.</summary>
internal static class GroupByFunctions
{
    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("PERCENTOF", 2, 2, [ArgumentKind.Any, ArgumentKind.Any], PercentOf));
        registry.Add(new FunctionInfo("GROUPBY", 3, 8,
            [ArgumentKind.ArrayContext, ArgumentKind.ArrayContext, ArgumentKind.Any, ArgumentKind.Value, ArgumentKind.Value,
             ArgumentKind.Any, ArgumentKind.ArrayContext, ArgumentKind.Value], Grouping.GroupBy));
        registry.Add(new FunctionInfo("PIVOTBY", 4, 11,
            [ArgumentKind.ArrayContext, ArgumentKind.ArrayContext, ArgumentKind.ArrayContext, ArgumentKind.Any, ArgumentKind.Value,
             ArgumentKind.Value, ArgumentKind.Any, ArgumentKind.Value, ArgumentKind.Any, ArgumentKind.ArrayContext, ArgumentKind.Value],
            Grouping.PivotBy)
        {
            Status = FunctionStatus.KnownDeviation,
            Deviation = "Several functions at once, such as HSTACK(SUM,AVERAGE), give #VALUE! instead of a column for each.",
        });
    }

    // PERCENTOF(data_subset, data_all): SUM(data_subset) / SUM(data_all). An error inside either
    // gives #NUM!, not the error itself (Excel for the web).
    private static Operand PercentOf(FunctionCall call)
    {
        var part = SumOf(call, 0, out var inData);
        if (part.IsError)
            return inData ? CellValue.Error(ErrorKind.Num) : part;
        var all = SumOf(call, 1, out inData);
        if (all.IsError)
            return inData ? CellValue.Error(ErrorKind.Num) : all;
        return all.AsNumber() == 0 ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(part.AsNumber() / all.AsNumber());
    }

    /// <summary>The sum of one argument as SUM takes it: numbers of ranges and arrays, a direct value converted.</summary>
    /// <param name="call">The call.</param>
    /// <param name="index">The argument.</param>
    /// <param name="inData">Whether an error result is an error value found in the data, rather than a value that could not be converted.</param>
    public static CellValue SumOf(FunctionCall call, int index, out bool inData)
    {
        var total = 0.0;
        var found = false;
        CellValue? error = null;
        Aggregation.ForEachIn(call, index, (value, source) =>
        {
            if (value.IsError)
            {
                error = value;
                found = true;
                return false;
            }

            if (value.Kind == CellValueKind.Number)
            {
                total += value.AsNumber();
                return true;
            }

            if (source != ValueSource.Direct || value.Kind is CellValueKind.Empty or CellValueKind.Missing)
                return true;
            var number = Coercion.ToNumber(value, call.Context.Culture, call.Context.DateSystem);
            if (number.IsError)
            {
                error = number;
                return false;
            }

            total += number.AsNumber();
            return true;
        });
        inData = found;
        return error ?? CellValue.Number(total);
    }
}
