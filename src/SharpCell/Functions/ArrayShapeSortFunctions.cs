using System;
using System.Collections.Generic;
using System.Globalization;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>SORT, SORTBY, FILTER, UNIQUE: dynamic array functions that pick and order rows (or columns).</summary>
internal static class ArrayShapeSortFunctions
{
    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("SORT", 1, 4, [ArgumentKind.ArrayContext, ArgumentKind.ArrayContext, ArgumentKind.ArrayContext, ArgumentKind.Value], Sort));
        registry.Add(new FunctionInfo("SORTBY", 2, FunctionRegistry.MaxArguments, [ArgumentKind.ArrayContext], SortBy));
        registry.Add(new FunctionInfo("FILTER", 2, 3, [ArgumentKind.ArrayContext, ArgumentKind.ArrayContext, ArgumentKind.Any], Filter));
        registry.Add(new FunctionInfo("UNIQUE", 1, 3, [ArgumentKind.ArrayContext, ArgumentKind.Value], Unique));
    }

    // Sorting order: numbers, text (ignoring case), booleans, errors (by code); descending
    // reverses that, but empty cells always come last. Equal keys keep their order.
    private static int CompareForSort(CellValue a, CellValue b, bool descending, CultureInfo culture)
    {
        var emptyA = a.Kind is CellValueKind.Empty or CellValueKind.Missing;
        var emptyB = b.Kind is CellValueKind.Empty or CellValueKind.Missing;
        if (emptyA || emptyB)
            return emptyA == emptyB ? 0 : emptyA ? 1 : -1;

        int order;
        if (a.IsError || b.IsError)
            order = a.IsError && b.IsError ? ((int)a.AsError()).CompareTo((int)b.AsError()) : a.IsError ? 1 : -1;
        else
            order = LookupSearch.Compare(a, b, culture);
        return descending ? -order : order;
    }

    // SORT(array, [index], [order], [by_col]): index and order may be arrays for several keys.
    private static Operand Sort(FunctionCall call)
    {
        if (!ArrayArguments.TryRead(call[0], call.Context, out var array, out var error))
            return error;
        var byColumn = call.Boolean(3, false);
        if (byColumn.IsError)
            return byColumn;

        var indexes = new List<long>();
        var orders = new List<long>();
        if (call.Has(1) && !ArrayArguments.TryReadIntegers(call, 1, indexes, out error))
            return error;
        if (call.Has(2) && !ArrayArguments.TryReadIntegers(call, 2, orders, out error))
            return error;
        if (indexes.Count == 0)
            indexes.Add(1);
        if (orders.Count > 1 && orders.Count != indexes.Count)
            return CellValue.Error(ErrorKind.Value);

        var columns = byColumn.AsBoolean();
        var width = array.GetLength(columns ? 0 : 1);
        var keys = new List<(Func<int, CellValue> Key, bool Descending)>();
        for (var i = 0; i < indexes.Count; i++)
        {
            var order = orders.Count == 0 ? 1 : orders[orders.Count == 1 ? 0 : i];
            var index = indexes[i];
            if (index < 1 || index > width || order is not (1 or -1))
                return CellValue.Error(ErrorKind.Value);
            var k = (int)index - 1;
            keys.Add((columns ? line => array[k, line] : line => array[line, k], order == -1));
        }

        return Reorder(array, columns, keys, call.Context);
    }

    // SORTBY(array, by, [order], by, [order], ...): each by array is a row or a column as long as
    // the array's side it sorts; all must sort the same side.
    private static Operand SortBy(FunctionCall call)
    {
        if (!ArrayArguments.TryRead(call[0], call.Context, out var array, out var error))
            return error;

        bool? columns = null;
        var keys = new List<(Func<int, CellValue> Key, bool Descending)>();
        for (var i = 1; i < call.Count; i += 2)
        {
            if (!ArrayArguments.TryRead(call[i], call.Context, out var by, out error))
                return error;
            var order = 1.0;
            if (i + 1 < call.Count && call.Has(i + 1))
            {
                var value = call.Value(i + 1);
                var number = Coercion.ToNumber(value.Kind == CellValueKind.Array ? value.AsArray()[0, 0] : value, call.Context.Culture, call.Context.DateSystem);
                if (number.IsError)
                    return number;
                order = Math.Truncate(number.AsNumber());
            }

            if (order is not (1 or -1))
                return CellValue.Error(ErrorKind.Value);

            // A column of the array's height sorts rows; a row of its width sorts columns.
            bool byColumns;
            if (by.GetLength(1) == 1 && by.GetLength(0) == array.GetLength(0))
                byColumns = false;
            else if (by.GetLength(0) == 1 && by.GetLength(1) == array.GetLength(1))
                byColumns = true;
            else
                return CellValue.Error(ErrorKind.Value);
            if (array.GetLength(0) == 1 && array.GetLength(1) == 1)
                byColumns = columns ?? false;
            if (columns is { } c && c != byColumns)
                return CellValue.Error(ErrorKind.Value);
            columns = byColumns;

            var key = by;
            keys.Add((byColumns ? line => key[0, line] : line => key[line, 0], order == -1));
        }

        return Reorder(array, columns ?? false, keys, call.Context);
    }

    private static Operand Reorder(CellValue[,] array, bool columns, List<(Func<int, CellValue> Key, bool Descending)> keys, EvaluationContext context)
    {
        var culture = context.Culture;
        var count = array.GetLength(columns ? 1 : 0);
        var lines = new int[count];
        for (var i = 0; i < count; i++)
            lines[i] = i;

        // A stable sort: ties fall back to the original position.
        Array.Sort(lines, (x, y) =>
        {
            foreach (var (key, descending) in keys)
            {
                var order = CompareForSort(key(x), key(y), descending, culture);
                if (order != 0)
                    return order;
            }

            return x.CompareTo(y);
        });
        context.CancellationToken.ThrowIfCancellationRequested();

        return CellValue.Array(Pick(array, columns, lines));
    }

    private static CellValue[,] Pick(CellValue[,] array, bool columns, IReadOnlyList<int> lines)
    {
        int rows = array.GetLength(0), width = array.GetLength(1);
        var result = columns ? new CellValue[rows, lines.Count] : new CellValue[lines.Count, width];
        for (var i = 0; i < lines.Count; i++)
        {
            if (columns)
            {
                for (var r = 0; r < rows; r++)
                    result[r, i] = array[r, lines[i]];
            }
            else
            {
                for (var c = 0; c < width; c++)
                    result[i, c] = array[lines[i], c];
            }
        }

        return result;
    }

    // FILTER(array, include, [if_empty]): the rows (include is a column of the array's height) or
    // columns (a row of its width) where include is TRUE or a non-zero number. Text in include is
    // #VALUE!, an error there is the result. Nothing kept: if_empty, or #CALC!.
    private static Operand Filter(FunctionCall call)
    {
        if (!ArrayArguments.TryRead(call[0], call.Context, out var array, out var error))
            return error;
        if (!ArrayArguments.TryRead(call[1], call.Context, out var include, out error))
            return error;

        int rows = array.GetLength(0), columns = array.GetLength(1);
        bool byColumns;
        if (include.GetLength(1) == 1 && include.GetLength(0) == rows)
            byColumns = false;
        else if (include.GetLength(0) == 1 && include.GetLength(1) == columns)
            byColumns = true;
        else
            return CellValue.Error(ErrorKind.Value);

        var kept = new List<int>();
        var count = byColumns ? columns : rows;
        for (var i = 0; i < count; i++)
        {
            var flag = byColumns ? include[0, i] : include[i, 0];
            switch (flag.Kind)
            {
                case CellValueKind.Error:
                    return flag;
                case CellValueKind.Text:
                    return CellValue.Error(ErrorKind.Value);
                case CellValueKind.Number when flag.AsNumber() != 0:
                case CellValueKind.Boolean when flag.AsBoolean():
                    kept.Add(i);
                    break;
            }
        }

        // if_empty is returned as a value, never as a reference: ISREF(FILTER(..., ..., A1)) is FALSE.
        if (kept.Count == 0)
            return call.Has(2) ? call.Value(2) : CellValue.Error(ErrorKind.Calc);
        return CellValue.Array(Pick(array, byColumns, kept));
    }

    // UNIQUE(array, [by_col], [exactly_once]): distinct rows (columns) in order of first appearance,
    // comparing text without case; with exactly_once only those that appear once.
    private static Operand Unique(FunctionCall call)
    {
        if (!ArrayArguments.TryRead(call[0], call.Context, out var array, out var error))
            return error;
        var byColumn = call.Boolean(1, false);
        if (byColumn.IsError)
            return byColumn;
        var exactlyOnce = call.Boolean(2, false);
        if (exactlyOnce.IsError)
            return exactlyOnce;

        var columns = byColumn.AsBoolean();
        var count = array.GetLength(columns ? 1 : 0);
        var comparer = new LineComparer(array, columns, call.Context.Culture);
        var occurrences = new Dictionary<int, int>(comparer);
        var order = new List<int>();
        for (var i = 0; i < count; i++)
        {
            if (occurrences.TryGetValue(i, out var seen))
            {
                occurrences[i] = seen + 1;
                continue;
            }

            occurrences.Add(i, 1);
            order.Add(i);
        }

        call.Context.CancellationToken.ThrowIfCancellationRequested();
        if (exactlyOnce.AsBoolean())
            order = order.FindAll(line => occurrences[line] == 1);
        if (order.Count == 0)
            return CellValue.Error(ErrorKind.Calc);
        return CellValue.Array(Pick(array, columns, order));
    }

    // Rows (or columns) of one array, compared element by element: same type and equal, text
    // without case. Hashes use the upper-case text and the exact number.
    private sealed class LineComparer(CellValue[,] array, bool columns, CultureInfo culture) : IEqualityComparer<int>
    {
        private readonly int _width = array.GetLength(columns ? 0 : 1);

        public bool Equals(int x, int y)
        {
            for (var i = 0; i < _width; i++)
            {
                var a = At(x, i);
                var b = At(y, i);
                if (a.Kind != b.Kind)
                    return false;
                var equal = a.Kind switch
                {
                    CellValueKind.Number => a.AsNumber() == b.AsNumber(),
                    CellValueKind.Text => TextComparer.AreEqual(a.AsText(), b.AsText(), culture),
                    CellValueKind.Boolean => a.AsBoolean() == b.AsBoolean(),
                    CellValueKind.Error => a.AsError() == b.AsError(),
                    _ => true,
                };
                if (!equal)
                    return false;
            }

            return true;
        }

        public int GetHashCode(int line)
        {
            var hash = new HashCode();
            for (var i = 0; i < _width; i++)
            {
                var value = At(line, i);
                hash.Add(value.Kind);
                switch (value.Kind)
                {
                    case CellValueKind.Number:
                        hash.Add(value.AsNumber());
                        break;
                    case CellValueKind.Text:
                        hash.Add(culture.TextInfo.ToUpper(value.AsText()), StringComparer.Ordinal);
                        break;
                    case CellValueKind.Boolean:
                        hash.Add(value.AsBoolean());
                        break;
                    case CellValueKind.Error:
                        hash.Add(value.AsError());
                        break;
                }
            }

            return hash.ToHashCode();
        }

        private CellValue At(int line, int i) => columns ? array[i, line] : array[line, i];
    }
}
