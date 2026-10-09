using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>Functions of whole arrays: SUMPRODUCT, the SUMX* pairs, matrix algebra, SEQUENCE and RANDARRAY.</summary>
internal static class MathArrayFunctions
{
    private const int CancellationCheckInterval = 4096;

    private static readonly ArgumentKind[] ArrayArguments = [ArgumentKind.ArrayContext];
    private static readonly ArgumentKind[] ValueArguments = [ArgumentKind.Value];

    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("SUMPRODUCT", 1, FunctionRegistry.MaxArguments, ArrayArguments, SumProduct));
        registry.Add(new FunctionInfo("SUMX2MY2", 2, 2, ArrayArguments, call => SumPairs(call, (x, y) => x * x - y * y)));
        registry.Add(new FunctionInfo("SUMX2PY2", 2, 2, ArrayArguments, call => SumPairs(call, (x, y) => x * x + y * y)));
        registry.Add(new FunctionInfo("SUMXMY2", 2, 2, ArrayArguments, call => SumPairs(call, (x, y) => (x - y) * (x - y))));
        registry.Add(new FunctionInfo("MMULT", 2, 2, ArrayArguments, MMult));
        registry.Add(new FunctionInfo("MDETERM", 1, 1, ArrayArguments, MDeterm));
        registry.Add(new FunctionInfo("MINVERSE", 1, 1, ArrayArguments, MInverse));
        registry.Add(new FunctionInfo("MUNIT", 1, 1, ValueArguments, MUnit));
        registry.Add(new FunctionInfo("SEQUENCE", 1, 4, ValueArguments, Sequence));
        registry.Add(new FunctionInfo("RANDARRAY", 0, 5, ValueArguments, RandArray) { IsVolatile = true });
    }

    // Arrays of one shape multiplied element by element and summed. Elements that are not numbers
    // count as 0; an error anywhere is the result. Only positions where some range holds a cell
    // are visited, so whole columns cost what they contain.
    private static Operand SumProduct(FunctionCall call)
    {
        var grids = new ValueGrid[call.Count];
        for (var i = 0; i < call.Count; i++)
        {
            if (call.IsMissing(i))
                return CellValue.Error(ErrorKind.Value);
            if (!ValueGrid.TryCreate(call[i], out grids[i], out var error))
                return error;
            if (grids[i].Rows != grids[0].Rows || grids[i].Columns != grids[0].Columns)
                return CellValue.Error(ErrorKind.Value);
        }

        var context = call.Context;
        var total = 0.0;
        var visited = 0;
        foreach (var (row, column) in Occupied(grids))
        {
            if (++visited % CancellationCheckInterval == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            var product = 1.0;
            foreach (var grid in grids)
            {
                var value = grid.Get(row, column, context);
                if (value.IsError)
                    return value;
                product = value.Kind == CellValueKind.Number ? product * value.AsNumber() : 0;
            }

            total += product;
        }

        return CellValue.Number(total);
    }

    // Elements are paired in reading order, so the two arrays need the same number of elements,
    // not the same shape. Pairs where either side is not a number are skipped.
    private static Operand SumPairs(FunctionCall call, Func<double, double, double> term)
    {
        if (call.IsMissing(0) || call.IsMissing(1))
            return CellValue.Error(ErrorKind.Value);
        if (!ValueGrid.TryCreate(call[0], out var x, out var error) || !ValueGrid.TryCreate(call[1], out var y, out error))
            return error;
        if ((long)x.Rows * x.Columns != (long)y.Rows * y.Columns)
            return CellValue.Error(ErrorKind.NA);

        // Positions of either array, as indexes in reading order.
        var indexes = new SortedSet<long>();
        AddIndexes(x, indexes);
        AddIndexes(y, indexes);

        var context = call.Context;
        var total = 0.0;
        var visited = 0;
        foreach (var index in indexes)
        {
            if (++visited % CancellationCheckInterval == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            var a = x.Get((int)(index / x.Columns), (int)(index % x.Columns), context);
            if (a.IsError)
                return a;
            var b = y.Get((int)(index / y.Columns), (int)(index % y.Columns), context);
            if (b.IsError)
                return b;
            if (a.Kind == CellValueKind.Number && b.Kind == CellValueKind.Number)
                total += term(a.AsNumber(), b.AsNumber());
        }

        return CellValue.Number(total);
    }

    private static void AddIndexes(ValueGrid grid, SortedSet<long> indexes)
    {
        var positions = new HashSet<(int, int)>();
        grid.AddOccupied(positions);
        foreach (var (row, column) in positions)
            indexes.Add((long)row * grid.Columns + column);
    }

    // Row by row, the order in which Excel meets errors and adds terms.
    private static List<(int Row, int Column)> Occupied(ValueGrid[] grids)
    {
        var occupied = new HashSet<(int, int)>();
        foreach (var grid in grids)
            grid.AddOccupied(occupied);
        var positions = new List<(int Row, int Column)>(occupied);
        positions.Sort();
        return positions;
    }

    private static Operand MMult(FunctionCall call)
    {
        if (!TryShape(call, 0, out var rows, out var inner, out var error) || !TryShape(call, 1, out var height, out var columns, out error))
            return error;
        if (inner != height)
            return CellValue.Error(ErrorKind.Value);
        if ((long)rows * columns > Evaluator.MaxArrayCells)
            return CellValue.Error(ErrorKind.Num);
        if (!TryReadMatrix(call, 0, out var a, out error) || !TryReadMatrix(call, 1, out var b, out error))
            return error;

        var result = new CellValue[rows, columns];
        for (var r = 0; r < rows; r++)
        {
            call.Context.CancellationToken.ThrowIfCancellationRequested();
            for (var c = 0; c < columns; c++)
            {
                var sum = 0.0;
                for (var k = 0; k < inner; k++)
                    sum += a[r, k] * b[k, c];
                result[r, c] = CellValue.Number(sum);
            }
        }

        return CellValue.Array(result);
    }

    // The product of the pivots of an LU decomposition, negated for each row exchange.
    private static Operand MDeterm(FunctionCall call)
    {
        if (!TryShape(call, 0, out var n, out var columns, out var error))
            return error;
        if (n != columns)
            return CellValue.Error(ErrorKind.Value);
        if (!TryReadMatrix(call, 0, out var m, out error))
            return error;
        if (!Decompose(m, call.Context, out var rows, out var exchanges))
            return CellValue.Number(0);

        var determinant = exchanges % 2 == 0 ? 1.0 : -1.0;
        for (var k = 0; k < rows.Length; k++)
            determinant *= m[k, k];
        return CellValue.Number(determinant);
    }

    // Each column of the inverse solves LU x = e_j by substitution, which reproduces Excel's
    // rounding (its MINVERSE has the same tiny nonzero entries); a singular matrix is #NUM!.
    private static Operand MInverse(FunctionCall call)
    {
        if (!TryShape(call, 0, out var n, out var columns, out var error))
            return error;
        if (n != columns)
            return CellValue.Error(ErrorKind.Value);
        if (!TryReadMatrix(call, 0, out var m, out error))
            return error;
        if (!Decompose(m, call.Context, out var rows, out _))
            return CellValue.Error(ErrorKind.Num);

        var result = new CellValue[n, n];
        var x = new double[n];
        for (var j = 0; j < n; j++)
        {
            call.Context.CancellationToken.ThrowIfCancellationRequested();
            for (var i = 0; i < n; i++)
            {
                var sum = 0.0;
                for (var k = 0; k < i; k++)
                    sum += m[i, k] * x[k];
                x[i] = (rows[i] == j ? 1 : 0) - sum;
            }

            for (var i = n - 1; i >= 0; i--)
            {
                var sum = 0.0;
                for (var k = i + 1; k < n; k++)
                    sum += m[i, k] * x[k];
                x[i] = (x[i] - sum) / m[i, i];
            }

            for (var i = 0; i < n; i++)
                result[i, j] = CellValue.Number(x[i]);
        }

        return CellValue.Array(result);
    }

    /// <summary>
    /// LU decomposition with partial pivoting, in place: below the diagonal the multipliers of L,
    /// from it up U. <paramref name="rows"/> maps each row of the result to its original row.
    /// </summary>
    /// <returns>False for a singular matrix.</returns>
    private static bool Decompose(double[,] m, EvaluationContext context, out int[] rows, out int exchanges)
    {
        var n = m.GetLength(0);
        rows = new int[n];
        for (var i = 0; i < n; i++)
            rows[i] = i;
        exchanges = 0;

        for (var k = 0; k < n; k++)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            var pivot = k;
            for (var r = k + 1; r < n; r++)
            {
                if (Math.Abs(m[r, k]) > Math.Abs(m[pivot, k]))
                    pivot = r;
            }

            if (m[pivot, k] == 0)
                return false;
            if (pivot != k)
            {
                for (var c = 0; c < n; c++)
                    (m[k, c], m[pivot, c]) = (m[pivot, c], m[k, c]);
                (rows[k], rows[pivot]) = (rows[pivot], rows[k]);
                exchanges++;
            }

            for (var r = k + 1; r < n; r++)
            {
                m[r, k] /= m[k, k];
                for (var c = k + 1; c < n; c++)
                    m[r, c] -= m[r, k] * m[k, c];
            }
        }

        return true;
    }

    // The shape of a matrix argument, known before any of its cells is read, so that a whole
    // column given to MDETERM is refused without reading a million cells.
    private static bool TryShape(FunctionCall call, int index, out int rows, out int columns, out CellValue error)
    {
        rows = columns = 0;
        error = CellValue.Error(ErrorKind.Value);
        if (call.IsMissing(index) || !ValueGrid.TryCreate(call[index], out var grid, out error))
            return false;
        rows = grid.Rows;
        columns = grid.Columns;
        return true;
    }

    /// <summary>
    /// The numbers of a matrix argument. Every element must be a number: the first one that is
    /// not decides the result, its own error or <c>#VALUE!</c> (an empty cell included).
    /// </summary>
    private static bool TryReadMatrix(FunctionCall call, int index, out double[,] matrix, out CellValue error)
    {
        matrix = new double[0, 0];
        error = CellValue.Error(ErrorKind.Value);
        if (call.IsMissing(index))
            return false;

        var value = call.Value(index);
        var array = value.Kind == CellValueKind.Array ? value.AsArray() : new[,] { { value } };
        matrix = new double[array.GetLength(0), array.GetLength(1)];
        for (var r = 0; r < array.GetLength(0); r++)
        {
            for (var c = 0; c < array.GetLength(1); c++)
            {
                var element = array[r, c];
                if (element.Kind != CellValueKind.Number)
                {
                    if (element.IsError)
                        error = element;
                    return false;
                }

                matrix[r, c] = element.AsNumber();
            }
        }

        return true;
    }

    // The n x n identity matrix.
    private static Operand MUnit(FunctionCall call)
    {
        var size = call.Integer(0);
        if (size.IsError)
            return size;
        var n = size.AsNumber();
        if (n < 1)
            return CellValue.Error(ErrorKind.Value);
        if (n * n > Evaluator.MaxArrayCells)
            return CellValue.Error(ErrorKind.Num);

        var result = new CellValue[(int)n, (int)n];
        for (var r = 0; r < n; r++)
        {
            for (var c = 0; c < n; c++)
                result[r, c] = CellValue.Number(r == c ? 1 : 0);
        }

        return CellValue.Array(result);
    }

    // rows x columns numbers from start by step, filled row by row. Each value adds the step to
    // the previous one, as Excel does: SEQUENCE(1,4,12.7,0.1) ends in 12.999999999999998.
    private static Operand Sequence(FunctionCall call)
    {
        if (!TryReadShape(call, 0, 1, out var rows, out var columns, out var error))
            return error;
        var start = call.Number(2, 1);
        if (start.IsError)
            return start;
        var step = call.Number(3, 1);
        if (step.IsError)
            return step;

        var result = new CellValue[rows, columns];
        var value = start.AsNumber();
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                result[r, c] = CellValue.Number(value);
                value += step.AsNumber();
            }
        }

        return CellValue.Array(result);
    }

    // Random numbers in [min, max), or whole numbers from CEILING(min) to FLOOR(max).
    private static Operand RandArray(FunctionCall call)
    {
        if (!TryReadShape(call, 0, 1, out var rows, out var columns, out var error))
            return error;
        var min = call.Number(2, 0);
        if (min.IsError)
            return min;
        var max = call.Number(3, 1);
        if (max.IsError)
            return max;
        var whole = call.Boolean(4, false);
        if (whole.IsError)
            return whole;

        double low = min.AsNumber(), high = max.AsNumber();
        if (whole.AsBoolean())
        {
            low = Math.Ceiling(low);
            high = Math.Floor(high);
        }

        if (low > high)
            return CellValue.Error(ErrorKind.Value);

        var random = call.Context.Workbook.Random;
        var result = new CellValue[rows, columns];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                var x = random.NextDouble();
                result[r, c] = CellValue.Number(whole.AsBoolean()
                    ? Math.Min(high, low + Math.Floor(x * (high - low + 1)))
                    : low + x * (high - low));
            }
        }

        return CellValue.Array(result);
    }

    // Row and column counts at arguments first and second (1 when absent or left empty); below 1 is #VALUE!.
    private static bool TryReadShape(FunctionCall call, int first, int second, out int rows, out int columns, out CellValue error)
    {
        rows = columns = 0;
        var r = call.Integer(first, 1);
        var c = call.Integer(second, 1);
        error = r.IsError ? r : c;
        if (r.IsError || c.IsError)
            return false;

        // A negative size is an invalid argument; a size of zero is an empty result (#CALC!), as in
        // Excel: SEQUENCE(-1) is #VALUE!, SEQUENCE(0.3) is #CALC!.
        error = CellValue.Error(ErrorKind.Value);
        if (r.AsNumber() < 0 || c.AsNumber() < 0)
            return false;
        error = CellValue.Error(ErrorKind.Calc);
        if (r.AsNumber() < 1 || c.AsNumber() < 1)
            return false;
        error = CellValue.Error(ErrorKind.Num);
        if (r.AsNumber() * c.AsNumber() > Evaluator.MaxArrayCells)
            return false;

        rows = (int)r.AsNumber();
        columns = (int)c.AsNumber();
        return true;
    }
}
