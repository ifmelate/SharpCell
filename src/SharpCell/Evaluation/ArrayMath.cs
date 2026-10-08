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

    /// <summary>
    /// Combines two values. The result has the larger height and width; a single row or column is
    /// repeated to fill it, and positions outside a smaller array become <c>#N/A</c>.
    /// </summary>
    public static CellValue Map(CellValue left, CellValue right, Func<CellValue, CellValue, CellValue> f)
    {
        if (left.Kind != CellValueKind.Array && right.Kind != CellValueKind.Array)
            return f(left, right);

        var a = AsArray(left);
        var b = AsArray(right);
        var rows = Math.Max(a.GetLength(0), b.GetLength(0));
        var columns = Math.Max(a.GetLength(1), b.GetLength(1));
        var result = new CellValue[rows, columns];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                var x = At(a, r, c);
                var y = At(b, r, c);
                result[r, c] = x is { } xv && y is { } yv ? f(xv, yv) : CellValue.Error(ErrorKind.NA);
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
