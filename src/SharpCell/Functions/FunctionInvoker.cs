using System;
using System.Collections.Generic;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Functions;

/// <summary>Calls a registry function: argument checks, element-wise application, the exception boundary.</summary>
internal static class FunctionInvoker
{
    public static Operand Invoke(FunctionNode node, EvaluationContext context)
    {
        if (!context.Workbook.Functions.TryGet(node.Name, out var function) || function!.Status == FunctionStatus.NotImplemented)
            return CellValue.Error(ErrorKind.Name);

        var arguments = node.Arguments;
        if (arguments.Count < function.MinArguments || arguments.Count > function.MaxArguments)
            return CellValue.Error(ErrorKind.Value);
        if (function.IsVolatile)
            context.UsedVolatile = true;

        var call = new FunctionCall(context, arguments);
        List<int>? lifted = null;
        for (var i = 0; i < arguments.Count; i++)
        {
            switch (function.KindAt(i))
            {
                case ArgumentKind.Any:
                    call.Set(i, Evaluator.Evaluate(arguments[i], context));
                    break;
                case ArgumentKind.Value:
                    var value = Evaluator.ToValue(Evaluator.Evaluate(arguments[i], context), context);
                    call.Set(i, value);
                    if (value.Kind == CellValueKind.Array)
                        (lifted ??= []).Add(i);
                    break;
            }
        }

        return lifted is null ? Run(function, call) : RunElementWise(function, call, lifted);
    }

    // A scalar parameter given an array: call once per element. Arrays broadcast as in operators,
    // and an element-wise result that is itself an array contributes its matching element.
    private static Operand RunElementWise(FunctionInfo function, FunctionCall call, List<int> lifted)
    {
        var arrays = new CellValue[lifted.Count][,];
        int rows = 1, columns = 1;
        for (var i = 0; i < lifted.Count; i++)
        {
            arrays[i] = call[lifted[i]].Value.AsArray();
            rows = Math.Max(rows, arrays[i].GetLength(0));
            columns = Math.Max(columns, arrays[i].GetLength(1));
        }

        var result = new CellValue[rows, columns];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                var inRange = true;
                for (var i = 0; i < lifted.Count && inRange; i++)
                {
                    if (TryAt(arrays[i], r, c, out var element))
                        call.Set(lifted[i], element);
                    else
                        inRange = false;
                }

                if (!inRange)
                {
                    result[r, c] = CellValue.Error(ErrorKind.NA);
                    continue;
                }

                var value = Evaluator.ToValue(Run(function, call), call.Context);
                result[r, c] = value.Kind != CellValueKind.Array
                    ? value
                    : TryAt(value.AsArray(), r, c, out var inner) ? inner : CellValue.Error(ErrorKind.NA);
            }
        }

        return CellValue.Array(result);
    }

    private static bool TryAt(CellValue[,] array, int row, int column, out CellValue value)
    {
        var rows = array.GetLength(0);
        var columns = array.GetLength(1);
        if ((rows != 1 && row >= rows) || (columns != 1 && column >= columns))
        {
            value = default;
            return false;
        }

        value = array[rows == 1 ? 0 : row, columns == 1 ? 0 : column];
        return true;
    }

    // A bug in a function must not take the calculation down: the cell gets #VALUE! and the
    // failure is reported. Cancellation is the one exception that passes through.
    private static Operand Run(FunctionInfo function, FunctionCall call)
    {
        try
        {
            return function.Body(call);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            call.Context.Report(DiagnosticKind.FunctionFailure, $"{function.Name} failed: {ex.GetType().Name}: {ex.Message}");
            return CellValue.Error(ErrorKind.Value);
        }
    }
}
