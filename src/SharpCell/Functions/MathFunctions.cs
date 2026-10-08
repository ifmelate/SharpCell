using System;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

internal static class MathFunctions
{
    private static readonly ArgumentKind[] AnyArguments = [ArgumentKind.Any];
    private static readonly ArgumentKind[] ValueArgument = [ArgumentKind.Value];

    public static void Register(FunctionRegistry registry)
    {
        var max = FunctionRegistry.MaxArguments;
        registry.Add(new FunctionInfo("SUM", 1, max, AnyArguments, Sum));
        registry.Add(new FunctionInfo("AVERAGE", 1, max, AnyArguments, Average));
        registry.Add(new FunctionInfo("MIN", 1, max, AnyArguments, call => Extreme(call, Math.Min)));
        registry.Add(new FunctionInfo("MAX", 1, max, AnyArguments, call => Extreme(call, Math.Max)));
        registry.Add(new FunctionInfo("COUNT", 1, max, AnyArguments, Count));
        registry.Add(new FunctionInfo("COUNTA", 1, max, AnyArguments, CountA));
        registry.Add(new FunctionInfo("ABS", 1, 1, ValueArgument, Abs));
        registry.Add(new FunctionInfo("RAND", 0, 0, ValueArgument, call => CellValue.Number(call.Context.Workbook.Random.NextDouble()))
        {
            IsVolatile = true,
        });
    }

    private static Operand Sum(FunctionCall call)
    {
        var total = 0.0;
        var error = Aggregation.Numbers(call, n => total += n);
        return error.IsError ? error : CellValue.Number(total);
    }

    private static Operand Average(FunctionCall call)
    {
        var total = 0.0;
        var count = 0;
        var error = Aggregation.Numbers(call, n =>
        {
            total += n;
            count++;
        });
        if (error.IsError)
            return error;
        return count == 0 ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(total / count);
    }

    // MIN and MAX of no numbers is 0.
    private static Operand Extreme(FunctionCall call, Func<double, double, double> pick)
    {
        double? result = null;
        var error = Aggregation.Numbers(call, n => result = result is { } r ? pick(r, n) : n);
        return error.IsError ? error : CellValue.Number(result ?? 0);
    }

    // Ranges and arrays: numbers only. Direct arguments: anything that converts to a number.
    // Errors are never counted and never propagate.
    private static Operand Count(FunctionCall call)
    {
        var count = 0;
        Aggregation.ForEach(call, (value, source) =>
        {
            if (value.Kind == CellValueKind.Number)
                count++;
            else if (source == ValueSource.Direct && !value.IsError && value.Kind != CellValueKind.Empty
                     && !Coercion.ToNumber(value, call.Context.Culture).IsError)
                count++;
            return true;
        });
        return CellValue.Number(count);
    }

    private static Operand CountA(FunctionCall call)
    {
        var count = 0;
        Aggregation.ForEach(call, (value, _) =>
        {
            if (value.Kind != CellValueKind.Empty)
                count++;
            return true;
        });
        return CellValue.Number(count);
    }

    private static Operand Abs(FunctionCall call)
    {
        var number = Coercion.ToNumber(call.Value(0), call.Context.Culture);
        return number.IsError ? number : CellValue.Number(Math.Abs(number.AsNumber()));
    }
}
