using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// Reading the array arguments of the dynamic array functions. A single-area reference is read
/// through the evaluation context (only stored cells; empty cells stay empty and spill as 0), a
/// scalar is a 1×1 array.
/// </summary>
internal static class ArrayArguments
{
    public static bool TryRead(Operand operand, EvaluationContext context, out CellValue[,] array, out CellValue error)
    {
        array = null!;
        if (!LookupTable.TryCreate(operand, out var table, out error))
            return false;
        if ((long)table.Rows * table.Columns > Evaluator.MaxArrayCells)
        {
            error = CellValue.Error(ErrorKind.Num);
            return false;
        }

        array = table.ToArray(0, 0, table.Rows, table.Columns, context);
        return true;
    }

    /// <summary>A result of rows × columns, or null when it is larger than Excel's limits allow in memory.</summary>
    public static CellValue[,]? Allocate(long rows, long columns) =>
        rows * columns > Evaluator.MaxArrayCells ? null : new CellValue[rows, columns];

    /// <summary>The numbers of an argument that may be a scalar or an array, row by row, truncated toward zero.</summary>
    public static bool TryReadIntegers(FunctionCall call, int index, List<long> numbers, out CellValue error)
    {
        error = default;
        var value = call.Value(index);
        var elements = value.Kind == CellValueKind.Array ? value.AsArray() : new[,] { { value } };
        foreach (var element in elements)
        {
            var number = Coercion.ToNumber(element, call.Context.Culture, call.Context.DateSystem);
            if (number.IsError)
            {
                error = number;
                return false;
            }

            numbers.Add((long)Math.Truncate(Math.Clamp(number.AsNumber(), -1e15, 1e15)));
        }

        return true;
    }
}

/// <summary>
/// Dynamic array functions that reshape arrays: TRANSPOSE, TAKE, DROP, CHOOSEROWS, CHOOSECOLS,
/// EXPAND, HSTACK, VSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS. An empty result is <c>#CALC!</c>.
/// </summary>
internal static class ArrayShapeFunctions
{
    private const int CancellationCheckInterval = 4096;

    public static void Register(FunctionRegistry registry)
    {
        var max = FunctionRegistry.MaxArguments;
        registry.Add(new FunctionInfo("TRANSPOSE", 1, 1, [ArgumentKind.ArrayContext], Transpose));
        registry.Add(new FunctionInfo("TAKE", 2, 3, [ArgumentKind.ArrayContext, ArgumentKind.Value], call => TakeOrDrop(call, take: true)));
        registry.Add(new FunctionInfo("DROP", 2, 3, [ArgumentKind.ArrayContext, ArgumentKind.Value], call => TakeOrDrop(call, take: false)));
        registry.Add(new FunctionInfo("CHOOSEROWS", 2, max, [ArgumentKind.ArrayContext, ArgumentKind.Any], call => Choose(call, rows: true)));
        registry.Add(new FunctionInfo("CHOOSECOLS", 2, max, [ArgumentKind.ArrayContext, ArgumentKind.Any], call => Choose(call, rows: false)));
        registry.Add(new FunctionInfo("EXPAND", 2, 4, [ArgumentKind.ArrayContext, ArgumentKind.Value], Expand));
        registry.Add(new FunctionInfo("HSTACK", 1, max, [ArgumentKind.ArrayContext], call => Stack(call, horizontal: true)));
        registry.Add(new FunctionInfo("VSTACK", 1, max, [ArgumentKind.ArrayContext], call => Stack(call, horizontal: false)));
        registry.Add(new FunctionInfo("TOCOL", 1, 3, [ArgumentKind.ArrayContext, ArgumentKind.Value], call => Flatten(call, column: true)));
        registry.Add(new FunctionInfo("TOROW", 1, 3, [ArgumentKind.ArrayContext, ArgumentKind.Value], call => Flatten(call, column: false)));
        registry.Add(new FunctionInfo("WRAPROWS", 2, 3, [ArgumentKind.ArrayContext, ArgumentKind.Value], call => Wrap(call, rows: true)));
        registry.Add(new FunctionInfo("WRAPCOLS", 2, 3, [ArgumentKind.ArrayContext, ArgumentKind.Value], call => Wrap(call, rows: false)));
    }

    private static Operand Transpose(FunctionCall call)
    {
        if (!ArrayArguments.TryRead(call[0], call.Context, out var array, out var error))
            return error;

        var rows = array.GetLength(0);
        var columns = array.GetLength(1);
        var result = new CellValue[columns, rows];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
                result[c, r] = array[r, c];
        }

        return CellValue.Array(result);
    }

    // TAKE(array, rows, [columns]) keeps the first rows (or with a negative count the last ones);
    // more than there are keeps all, 0 keeps nothing. DROP(array, rows, [columns]) removes them.
    // Only the kept part of a reference is read.
    private static Operand TakeOrDrop(FunctionCall call, bool take)
    {
        if (!LookupTable.TryCreate(call[0], out var table, out var error))
            return error;
        if (!TryCount(call, 1, out var rows, out error) || !TryCount(call, 2, out var columns, out error))
            return error;

        var (firstRow, rowCount) = Keep(table.Rows, rows, take);
        var (firstColumn, columnCount) = Keep(table.Columns, columns, take);
        if (rowCount == 0 || columnCount == 0)
            return CellValue.Error(ErrorKind.Calc);
        if ((long)rowCount * columnCount > Evaluator.MaxArrayCells)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Array(table.ToArray(firstRow, firstColumn, rowCount, columnCount, call.Context));
    }

    // An absent count keeps the whole side.
    private static bool TryCount(FunctionCall call, int index, out long? count, out CellValue error)
    {
        count = null;
        error = default;
        if (!call.Has(index))
            return true;
        var number = call.Number(index);
        if (number.IsError)
        {
            error = number;
            return false;
        }

        count = (long)Math.Truncate(Math.Clamp(number.AsNumber(), -1e15, 1e15));
        return true;
    }

    private static (int First, int Count) Keep(int length, long? count, bool take)
    {
        if (count is not { } n)
            return (0, length);
        var magnitude = (int)Math.Min(Math.Abs(n), length);
        if (take)
            return n >= 0 ? (0, magnitude) : (length - magnitude, magnitude);
        return n >= 0 ? (magnitude, length - magnitude) : (0, length - magnitude);
    }

    // CHOOSEROWS(array, row, ...) and CHOOSECOLS: the listed rows (columns) in the listed order,
    // repeats allowed; a negative number counts from the end. 0 or beyond the array is #VALUE!.
    private static Operand Choose(FunctionCall call, bool rows)
    {
        if (!LookupTable.TryCreate(call[0], out var table, out var error))
            return error;

        var picks = new List<long>();
        for (var i = 1; i < call.Count; i++)
        {
            if (!ArrayArguments.TryReadIntegers(call, i, picks, out error))
                return error;
        }

        var length = rows ? table.Rows : table.Columns;
        var other = rows ? table.Columns : table.Rows;
        if (ArrayArguments.Allocate(rows ? picks.Count : other, rows ? other : picks.Count) is not { } result)
            return CellValue.Error(ErrorKind.Num);
        for (var p = 0; p < picks.Count; p++)
        {
            var pick = picks[p];
            if (pick == 0 || Math.Abs(pick) > length)
                return CellValue.Error(ErrorKind.Value);
            var index = (int)(pick > 0 ? pick - 1 : length + pick);
            for (var j = 0; j < other; j++)
            {
                if (rows)
                    result[p, j] = table.Get(index, j, call.Context);
                else
                    result[j, p] = table.Get(j, index, call.Context);
            }
        }

        return CellValue.Array(result);
    }

    // EXPAND(array, rows, [columns], [pad_with]): grows the array, filling with pad_with (#N/A by
    // default). A size below the array's is #VALUE!.
    private static Operand Expand(FunctionCall call)
    {
        if (!ArrayArguments.TryRead(call[0], call.Context, out var array, out var error))
            return error;
        int rows = array.GetLength(0), columns = array.GetLength(1);
        if (!TryCount(call, 1, out var height, out error) || !TryCount(call, 2, out var width, out error))
            return error;
        var pad = call.Has(3) ? call.Value(3) : CellValue.Error(ErrorKind.NA);

        var newRows = height ?? rows;
        var newColumns = width ?? columns;
        if (newRows < rows || newColumns < columns)
            return CellValue.Error(ErrorKind.Value);
        if (newRows > CellAddress.MaxRow || newColumns > CellAddress.MaxColumn || ArrayArguments.Allocate(newRows, newColumns) is not { } result)
            return CellValue.Error(ErrorKind.Num);

        for (var r = 0; r < newRows; r++)
        {
            for (var c = 0; c < newColumns; c++)
                result[r, c] = r < rows && c < columns ? array[r, c] : pad;
        }

        return CellValue.Array(result);
    }

    // HSTACK side by side, VSTACK one under another; shorter arrays are padded with #N/A.
    private static Operand Stack(FunctionCall call, bool horizontal)
    {
        var parts = new List<CellValue[,]>();
        long along = 0, across = 0;
        for (var i = 0; i < call.Count; i++)
        {
            if (call.IsMissing(i))
                continue;
            if (!ArrayArguments.TryRead(call[i], call.Context, out var part, out var error))
                return error;
            parts.Add(part);
            along += part.GetLength(horizontal ? 1 : 0);
            across = Math.Max(across, part.GetLength(horizontal ? 0 : 1));
        }

        if (parts.Count == 0)
            return CellValue.Error(ErrorKind.Value);
        if (ArrayArguments.Allocate(horizontal ? across : along, horizontal ? along : across) is not { } result)
            return CellValue.Error(ErrorKind.Num);

        var na = CellValue.Error(ErrorKind.NA);
        var offset = 0;
        foreach (var part in parts)
        {
            int rows = part.GetLength(0), columns = part.GetLength(1);
            var height = horizontal ? (int)across : rows;
            var width = horizontal ? columns : (int)across;
            for (var r = 0; r < height; r++)
            {
                for (var c = 0; c < width; c++)
                {
                    var value = r < rows && c < columns ? part[r, c] : na;
                    if (horizontal)
                        result[r, offset + c] = value;
                    else
                        result[offset + r, c] = value;
                }
            }

            offset += horizontal ? columns : rows;
        }

        return CellValue.Array(result);
    }

    // TOCOL(array, [ignore], [scan_by_column]) and TOROW: the elements in one column (row), read
    // row by row unless scan_by_column. ignore: 0 keeps all, 1 skips blanks, 2 errors, 3 both. A
    // reference with blanks skipped is walked by its stored cells, so TOCOL(A:A,1) is cheap.
    private static Operand Flatten(FunctionCall call, bool column)
    {
        var ignore = call.Integer(1, 0);
        if (ignore.IsError)
            return ignore;
        var byColumn = call.Boolean(2, false);
        if (byColumn.IsError)
            return byColumn;
        if (ignore.AsNumber() is not (0 or 1 or 2 or 3))
            return CellValue.Error(ErrorKind.Value);
        var skipBlanks = ignore.AsNumber() is 1 or 3;
        var skipErrors = ignore.AsNumber() is 2 or 3;

        if (!LookupTable.TryCreate(call[0], out var table, out var error))
            return error;

        var context = call.Context;
        var kept = new List<(int Row, int Column, CellValue Value)>();
        var visited = 0;
        if (skipBlanks)
        {
            foreach (var cell in table.Occupied(0, 0, table.Rows - 1, table.Columns - 1, context))
            {
                if (++visited % CancellationCheckInterval == 0)
                    context.CancellationToken.ThrowIfCancellationRequested();
                if (cell.Value.Kind != CellValueKind.Empty && !(skipErrors && cell.Value.IsError))
                    kept.Add(cell);
            }

            if (byColumn.AsBoolean())
                kept.Sort((a, b) => a.Column != b.Column ? a.Column.CompareTo(b.Column) : a.Row.CompareTo(b.Row));
        }
        else
        {
            if ((long)table.Rows * table.Columns > Evaluator.MaxArrayCells)
                return CellValue.Error(ErrorKind.Num);
            var array = table.ToArray(0, 0, table.Rows, table.Columns, context);
            int rows = table.Rows, columns = table.Columns;
            var outer = byColumn.AsBoolean() ? columns : rows;
            var inner = byColumn.AsBoolean() ? rows : columns;
            for (var i = 0; i < outer; i++)
            {
                for (var j = 0; j < inner; j++)
                {
                    var (r, c) = byColumn.AsBoolean() ? (j, i) : (i, j);
                    if (!(skipErrors && array[r, c].IsError))
                        kept.Add((r, c, array[r, c]));
                }
            }
        }

        if (kept.Count == 0)
            return CellValue.Error(ErrorKind.Calc);
        var result = column ? new CellValue[kept.Count, 1] : new CellValue[1, kept.Count];
        for (var i = 0; i < kept.Count; i++)
        {
            if (column)
                result[i, 0] = kept[i].Value;
            else
                result[0, i] = kept[i].Value;
        }

        return CellValue.Array(result);
    }

    // WRAPROWS(vector, count, [pad_with]) lays a row or column out in rows of count elements;
    // WRAPCOLS in columns. The last row (column) is padded with pad_with, #N/A by default.
    private static Operand Wrap(FunctionCall call, bool rows)
    {
        if (!ArrayArguments.TryRead(call[0], call.Context, out var array, out var error))
            return error;
        var count = call.Integer(1);
        if (count.IsError)
            return count;
        var pad = call.Has(2) ? call.Value(2) : CellValue.Error(ErrorKind.NA);
        if (array.GetLength(0) != 1 && array.GetLength(1) != 1)
            return CellValue.Error(ErrorKind.Value);
        if (count.AsNumber() < 1)
            return CellValue.Error(ErrorKind.Num);

        var length = array.Length;
        var size = (long)Math.Min(count.AsNumber(), CellAddress.MaxRow + 1.0);
        var lines = (length + size - 1) / size;
        if ((rows ? size > CellAddress.MaxColumn : size > CellAddress.MaxRow) || ArrayArguments.Allocate(lines, size) is null)
            return CellValue.Error(ErrorKind.Num);
        var result = rows ? new CellValue[lines, size] : new CellValue[size, lines];
        var horizontal = array.GetLength(0) == 1;
        for (var line = 0; line < lines; line++)
        {
            for (var i = 0; i < size; i++)
            {
                var position = (line * size) + i;
                var value = position < length ? (horizontal ? array[0, position] : array[position, 0]) : pad;
                if (rows)
                    result[line, i] = value;
                else
                    result[i, line] = value;
            }
        }

        return CellValue.Array(result);
    }
}
