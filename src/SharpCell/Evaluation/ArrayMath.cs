using System;

namespace SharpCell.Evaluation;

/// <summary>Element-wise application of scalar operations to arrays, with Excel's broadcasting.</summary>
internal static class ArrayMath
{
    public static CellValue Map(CellValue value, Func<CellValue, CellValue> f)
    {
        if (value.Kind != CellValueKind.Array)
            return f(value);

        var source = value.AsArray();
        var result = new CellValue[source.GetLength(0), source.GetLength(1)];
        for (var r = 0; r < source.GetLength(0); r++)
        {
            for (var c = 0; c < source.GetLength(1); c++)
                result[r, c] = f(source[r, c]);
        }

        return CellValue.Array(result);
    }

    private static readonly CellValue NA = CellValue.Error(ErrorKind.NA);

    /// <summary>
    /// Combines two values. The result has the larger height and width; a single row or column is
    /// repeated to fill it, and a smaller array is padded with <c>#N/A</c> before combining, as in
    /// Excel: <paramref name="f"/> sees the padding, so IFERROR can replace it and an error on the
    /// other side still wins (<c>{#DIV/0!;2;3}+{1;2}</c> gives #DIV/0!, 4, #N/A).
    /// </summary>
    public static CellValue Map(CellValue left, CellValue right, Func<CellValue, CellValue, CellValue> f)
    {
        if (left.Kind != CellValueKind.Array && right.Kind != CellValueKind.Array)
            return f(left, right);

        var a = AsArray(left);
        var b = AsArray(right);
        var rows = Math.Max(a.GetLength(0), b.GetLength(0));
        var columns = Math.Max(a.GetLength(1), b.GetLength(1));
        if ((long)rows * columns > Evaluator.MaxArrayCells)
            return CellValue.Error(ErrorKind.Num);
        var result = new CellValue[rows, columns];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                result[r, c] = f(At(a, r, c) ?? NA, At(b, r, c) ?? NA);
            }
        }

        return CellValue.Array(result);
    }

    /// <summary>Three-way element-wise combination with the same broadcasting rules.</summary>
    public static CellValue Map(CellValue first, CellValue second, CellValue third, Func<CellValue, CellValue, CellValue, CellValue> f)
    {
        if (first.Kind != CellValueKind.Array && second.Kind != CellValueKind.Array && third.Kind != CellValueKind.Array)
            return f(first, second, third);

        var a = AsArray(first);
        var b = AsArray(second);
        var c = AsArray(third);
        var rows = Math.Max(a.GetLength(0), Math.Max(b.GetLength(0), c.GetLength(0)));
        var columns = Math.Max(a.GetLength(1), Math.Max(b.GetLength(1), c.GetLength(1)));
        if ((long)rows * columns > Evaluator.MaxArrayCells)
            return CellValue.Error(ErrorKind.Num);
        var result = new CellValue[rows, columns];
        for (var r = 0; r < rows; r++)
        {
            for (var col = 0; col < columns; col++)
            {
                result[r, col] = f(At(a, r, col) ?? NA, At(b, r, col) ?? NA, At(c, r, col) ?? NA);
            }
        }

        return CellValue.Array(result);
    }

    private static CellValue[,] AsArray(CellValue value) =>
        value.Kind == CellValueKind.Array ? value.AsArray() : new[,] { { value } };

    private static CellValue? At(CellValue[,] array, int row, int column)
    {
        var rows = array.GetLength(0);
        var columns = array.GetLength(1);
        if ((rows != 1 && row >= rows) || (columns != 1 && column >= columns))
            return null;
        return array[rows == 1 ? 0 : row, columns == 1 ? 0 : column];
    }
}
