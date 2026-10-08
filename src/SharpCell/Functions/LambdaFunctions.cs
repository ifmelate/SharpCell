using System;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// Functions that call a LAMBDA per element, row or column. Inputs are read first; when that met a
/// dirty cell no lambda is called, since a stand-in value could steer the lambda into cells the real
/// evaluation never reads.
/// </summary>
internal static class LambdaFunctions
{
    private const int CancellationCheckInterval = 4096;

    public static void Register(FunctionRegistry registry)
    {
        ArgumentKind[] any = [ArgumentKind.Any];
        registry.Add(new FunctionInfo("MAP", 2, FunctionRegistry.MaxArguments, any, Map));
        registry.Add(new FunctionInfo("REDUCE", 3, 3, any, call => Fold(call, scan: false)));
        registry.Add(new FunctionInfo("SCAN", 3, 3, any, call => Fold(call, scan: true)));
        registry.Add(new FunctionInfo("BYROW", 2, 2, any, call => ByLine(call, rows: true)));
        registry.Add(new FunctionInfo("BYCOL", 2, 2, any, call => ByLine(call, rows: false)));
        registry.Add(new FunctionInfo("MAKEARRAY", 3, 3, any, MakeArray));
    }

    // MAP(array1, [array2, ...], lambda): arrays broadcast like operators; one parameter per array.
    private static Operand Map(FunctionCall call)
    {
        var count = call.Count - 1;
        if (!TryGetLambda(call, count, out var lambda, out var error))
            return error;
        if (lambda.Parameters.Length != count)
            return CellValue.Error(ErrorKind.Value);

        var arrays = new CellValue[count][,];
        int rows = 1, columns = 1;
        for (var i = 0; i < count; i++)
        {
            var value = call.Value(i);
            arrays[i] = value.Kind == CellValueKind.Array ? value.AsArray() : new[,] { { value } };
            rows = Math.Max(rows, arrays[i].GetLength(0));
            columns = Math.Max(columns, arrays[i].GetLength(1));
        }

        if (call.MetPendingInput)
            return EvaluationContext.PendingPlaceholder;
        if ((long)rows * columns > Evaluator.MaxArrayCells)
            return CellValue.Error(ErrorKind.Num);

        var result = new CellValue[rows, columns];
        var arguments = new Operand[count];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                var inRange = true;
                for (var i = 0; i < count && inRange; i++)
                {
                    if (At(arrays[i], r, c) is { } element)
                        arguments[i] = element;
                    else
                        inRange = false;
                }

                result[r, c] = inRange ? Element(lambda, arguments, call, r * columns + c) : CellValue.Error(ErrorKind.NA);
            }
        }

        return CellValue.Array(result);
    }

    // REDUCE([initial], array, lambda(acc, value)) and SCAN: row-major over the array. An omitted
    // initial value starts the accumulator as an omitted argument (0 in arithmetic).
    private static Operand Fold(FunctionCall call, bool scan)
    {
        if (!TryGetLambda(call, 2, out var lambda, out var error))
            return error;
        if (lambda.Parameters.Length != 2)
            return CellValue.Error(ErrorKind.Value);

        var accumulator = call.IsMissing(0) ? CellValue.Missing : call.Value(0);
        var input = call.Value(1);
        if (call.MetPendingInput)
            return EvaluationContext.PendingPlaceholder;

        var array = input.Kind == CellValueKind.Array ? input.AsArray() : new[,] { { input } };
        var steps = scan ? new CellValue[array.GetLength(0), array.GetLength(1)] : null;
        var index = 0;
        for (var r = 0; r < array.GetLength(0); r++)
        {
            for (var c = 0; c < array.GetLength(1); c++)
            {
                CheckCancellation(call, index++);
                accumulator = Evaluator.ToValue(Lambdas.Invoke(lambda, [accumulator, array[r, c]], call.Context), call.Context);
                if (steps is not null)
                    steps[r, c] = Scalar(accumulator);
            }
        }

        return steps is null ? accumulator : CellValue.Array(steps);
    }

    // BYROW/BYCOL(array, lambda(line)): one call per row or column, each giving one value. A range
    // passes each line as a reference, an array as a one-row or one-column array.
    private static Operand ByLine(FunctionCall call, bool rows)
    {
        if (!TryGetLambda(call, 1, out var lambda, out var error))
            return error;
        if (lambda.Parameters.Length != 1)
            return CellValue.Error(ErrorKind.Value);

        var input = call[0];
        Operand[] lines;
        if (input.Reference is { } reference)
        {
            if (!reference.IsSingleArea)
                return CellValue.Error(ErrorKind.Value);

            var (sheet, area) = reference.Areas[0];
            var count = rows ? area.Rows : area.Columns;
            if (count > Evaluator.MaxArrayCells)
                return CellValue.Error(ErrorKind.Num);

            lines = new Operand[count];
            for (var i = 0; i < count; i++)
            {
                var line = rows
                    ? new Area(area.FirstRow + i, area.FirstColumn, area.FirstRow + i, area.LastColumn)
                    : new Area(area.FirstRow, area.FirstColumn + i, area.LastRow, area.FirstColumn + i);
                lines[i] = Operand.Of(new Reference(sheet, line));
            }
        }
        else
        {
            var value = input.Value;
            if (value.IsError)
                return value;

            var array = value.Kind == CellValueKind.Array ? value.AsArray() : new[,] { { value } };
            var count = array.GetLength(rows ? 0 : 1);
            var length = array.GetLength(rows ? 1 : 0);
            lines = new Operand[count];
            for (var i = 0; i < count; i++)
            {
                var line = rows ? new CellValue[1, length] : new CellValue[length, 1];
                for (var j = 0; j < length; j++)
                {
                    if (rows)
                        line[0, j] = array[i, j];
                    else
                        line[j, 0] = array[j, i];
                }

                lines[i] = CellValue.Array(line);
            }
        }

        if (call.MetPendingInput)
            return EvaluationContext.PendingPlaceholder;

        var result = rows ? new CellValue[lines.Length, 1] : new CellValue[1, lines.Length];
        for (var i = 0; i < lines.Length; i++)
        {
            var value = Element(lambda, [lines[i]], call, i);
            if (rows)
                result[i, 0] = value;
            else
                result[0, i] = value;
        }

        return CellValue.Array(result);
    }

    // MAKEARRAY(rows, columns, lambda(row, column)) with 1-based indexes.
    private static Operand MakeArray(FunctionCall call)
    {
        if (!TryGetLambda(call, 2, out var lambda, out var error))
            return error;
        if (lambda.Parameters.Length != 2)
            return CellValue.Error(ErrorKind.Value);

        var rowsValue = Coercion.ToNumber(call.Value(0), call.Context.Culture);
        var columnsValue = Coercion.ToNumber(call.Value(1), call.Context.Culture);
        if (call.MetPendingInput)
            return EvaluationContext.PendingPlaceholder;
        if (rowsValue.IsError)
            return rowsValue;
        if (columnsValue.IsError)
            return columnsValue;

        var rows = Math.Truncate(rowsValue.AsNumber());
        var columns = Math.Truncate(columnsValue.AsNumber());
        if (rows < 1 || columns < 1)
            return CellValue.Error(ErrorKind.Value);
        if (rows * columns > Evaluator.MaxArrayCells)
            return CellValue.Error(ErrorKind.Num);

        var result = new CellValue[(int)rows, (int)columns];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
                result[r, c] = Element(lambda, [CellValue.Number(r + 1), CellValue.Number(c + 1)], call, (int)(r * columns + c));
        }

        return CellValue.Array(result);
    }

    private static bool TryGetLambda(FunctionCall call, int index, out Closure lambda, out CellValue error)
    {
        var value = call.Value(index);
        if (value.Kind == CellValueKind.Lambda && value.AsLambda() is Closure closure)
        {
            lambda = closure;
            error = default;
            return true;
        }

        lambda = null!;
        error = value.IsError ? value : CellValue.Error(ErrorKind.Value);
        return false;
    }

    private static CellValue Element(Closure lambda, Operand[] arguments, FunctionCall call, int index)
    {
        CheckCancellation(call, index);
        return Scalar(Evaluator.ToValue(Lambdas.Invoke(lambda, arguments, call.Context), call.Context));
    }

    // An element of a result cannot itself be an array or a function.
    private static CellValue Scalar(CellValue value) =>
        value.Kind is CellValueKind.Array or CellValueKind.Lambda ? CellValue.Error(ErrorKind.Calc) : value;

    private static void CheckCancellation(FunctionCall call, int index)
    {
        if (index % CancellationCheckInterval == CancellationCheckInterval - 1)
            call.Context.CancellationToken.ThrowIfCancellationRequested();
    }

    private static CellValue? At(CellValue[,] array, int row, int column)
    {
        var rows = array.GetLength(0);
        var columns = array.GetLength(1);
        if ((rows != 1 && row >= rows) || (columns != 1 && column >= columns))
            return null;
        return array[rows == 1 ? 0 : row, columns == 1 ? 0 : column];
    }
}
