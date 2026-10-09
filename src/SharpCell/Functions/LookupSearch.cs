using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// The range or array argument of a lookup function: a single-area reference, read cell by cell
/// through the evaluation context (dirty cells, dependencies), or an array. A scalar is a 1×1 array.
/// </summary>
internal readonly struct LookupTable
{
    private readonly CellValue[,]? _array;

    private LookupTable(Worksheet? sheet, Area area, CellValue[,]? array)
    {
        Sheet = sheet;
        Area = area;
        _array = array;
        Rows = array?.GetLength(0) ?? area.Rows;
        Columns = array?.GetLength(1) ?? area.Columns;
    }

    /// <summary>The sheet of a reference; null for an array.</summary>
    public Worksheet? Sheet { get; }

    /// <summary>The area of a reference.</summary>
    public Area Area { get; }

    public int Rows { get; }

    public int Columns { get; }

    public bool IsReference => Sheet is not null;

    public bool IsVector => Rows == 1 || Columns == 1;

    /// <summary>A reference with several areas is <c>#VALUE!</c>; an error is itself.</summary>
    public static bool TryCreate(Operand operand, out LookupTable table, out CellValue error)
    {
        error = default;
        table = default;
        if (operand.Reference is { } reference)
        {
            if (!reference.IsSingleArea)
            {
                error = CellValue.Error(ErrorKind.Value);
                return false;
            }

            table = new LookupTable(reference.Areas[0].Sheet, reference.Areas[0].Area, null);
            return true;
        }

        var value = operand.Value;
        if (value.IsError)
        {
            error = value;
            return false;
        }

        table = new LookupTable(null, default, value.Kind == CellValueKind.Array ? value.AsArray() : new[,] { { value } });
        return true;
    }

    public CellValue Get(int row, int column, EvaluationContext context) =>
        _array is not null ? _array[row, column] : context.ReadCell(Sheet!, Area.FirstRow + row, Area.FirstColumn + column);

    public LookupVector Row(int row) => new(this, row, horizontal: true);

    public LookupVector Column(int column) => new(this, column, horizontal: false);

    /// <summary>The vector a one-dimensional table is: its row if it has one row, otherwise its column.</summary>
    public LookupVector AsVector() => Rows == 1 ? Row(0) : Column(0);

    /// <summary>
    /// A rectangle of the table as a function result: a reference into a reference (recorded as a
    /// dependency), a sub-array of an array, a single value for one element of an array.
    /// </summary>
    public Operand Slice(int firstRow, int firstColumn, int rows, int columns, EvaluationContext context)
    {
        if (Sheet is { } sheet)
        {
            var area = new Area(Area.FirstRow + firstRow, Area.FirstColumn + firstColumn,
                Area.FirstRow + firstRow + rows - 1, Area.FirstColumn + firstColumn + columns - 1);
            var result = new Reference(sheet, area);
            context.RecordReference(result);
            return Operand.Of(result);
        }

        if (rows == 1 && columns == 1)
            return _array![firstRow, firstColumn];

        var values = new CellValue[rows, columns];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
                values[r, c] = _array![firstRow + r, firstColumn + c];
        }

        return CellValue.Array(values);
    }

    /// <summary>
    /// A rectangle of the table as a new array; for a reference only stored cells are read, the
    /// others stay empty. The caller keeps the size within <see cref="Evaluator.MaxArrayCells"/>.
    /// </summary>
    public CellValue[,] ToArray(int firstRow, int firstColumn, int rows, int columns, EvaluationContext context)
    {
        var values = new CellValue[rows, columns];
        var visited = 0;
        foreach (var (row, column, value) in Occupied(firstRow, firstColumn, firstRow + rows - 1, firstColumn + columns - 1, context))
        {
            if (++visited % 4096 == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            values[row - firstRow, column - firstColumn] = value;
        }

        return values;
    }

    /// <summary>Positions (row, column) in the rectangle that may hold a value, row by row; for a reference only stored cells.</summary>
    public IEnumerable<(int Row, int Column, CellValue Value)> Occupied(int firstRow, int firstColumn, int lastRow, int lastColumn, EvaluationContext context)
    {
        if (Sheet is not { } sheet)
        {
            for (var r = firstRow; r <= lastRow; r++)
            {
                for (var c = firstColumn; c <= lastColumn; c++)
                    yield return (r, c, _array![r, c]);
            }

            yield break;
        }

        foreach (var cell in sheet.Store.Enumerate(Area.FirstRow + firstRow, Area.FirstColumn + firstColumn, Area.FirstRow + lastRow, Area.FirstColumn + lastColumn))
            yield return (cell.Row - Area.FirstRow, cell.Column - Area.FirstColumn, context.ReadCell(sheet, cell));
    }
}

/// <summary>One row or column of a <see cref="LookupTable"/>, indexed from 0.</summary>
internal readonly struct LookupVector(LookupTable table, int index, bool horizontal)
{
    public int Length => horizontal ? table.Columns : table.Rows;

    public CellValue Get(int i, EvaluationContext context) => horizontal ? table.Get(index, i, context) : table.Get(i, index, context);

    /// <summary>Positions in [from, to] that may hold a value, ascending; for a reference only stored cells are visited.</summary>
    public IEnumerable<(int Index, CellValue Value)> Occupied(int from, int to, EvaluationContext context)
    {
        if (from > to)
            yield break;
        var cells = horizontal
            ? table.Occupied(index, from, index, to, context)
            : table.Occupied(from, index, to, index, context);
        foreach (var (row, column, value) in cells)
            yield return (horizontal ? column : row, value);
    }
}

/// <summary>
/// How MATCH, VLOOKUP, XLOOKUP and the others find a value. Two families:
/// <list type="bullet">
/// <item>The legacy functions (MATCH, VLOOKUP, HLOOKUP, LOOKUP) compare a value only with values of
/// its own type; an empty lookup value means 0. Approximate matches use a binary search that skips
/// values of other types, errors and empty cells, which is why <c>LOOKUP(2,1/(A:A="x"),B:B)</c>
/// finds the last match.</item>
/// <item>XLOOKUP and XMATCH order all types (numbers &lt; text &lt; booleans) and skip errors; an
/// empty lookup value finds an empty cell.</item>
/// </list>
/// Text compares ignoring case, numbers with Excel's tolerance for floating-point noise.
/// </summary>
internal static class LookupSearch
{
    private const int CancellationCheckInterval = 4096;

    public const int ExactMatch = 0;
    public const int ExactOrSmaller = -1;
    public const int ExactOrLarger = 1;
    public const int WildcardMatch = 2;
    public const int RegexMatch = 3;

    public const int FirstToLast = 1;
    public const int LastToFirst = -1;
    public const int BinaryAscending = 2;
    public const int BinaryDescending = -2;

    /// <summary>Type order of XLOOKUP and SORT: number, text, boolean; -1 for what lookups skip (errors, empty).</summary>
    public static int Rank(CellValue value) => value.Kind switch
    {
        CellValueKind.Number => 0,
        CellValueKind.Text => 1,
        CellValueKind.Boolean => 2,
        _ => -1,
    };

    /// <summary>Orders two values of known rank: by type first, then within the type.</summary>
    public static int Compare(CellValue a, CellValue b, CultureInfo culture)
    {
        var rankA = Rank(a);
        var rankB = Rank(b);
        if (rankA != rankB)
            return rankA.CompareTo(rankB);
        return a.Kind switch
        {
            CellValueKind.Number => NumberComparer.Compare(a.AsNumber(), b.AsNumber()),
            CellValueKind.Text => TextComparer.Compare(a.AsText(), b.AsText(), culture),
            CellValueKind.Boolean => a.AsBoolean().CompareTo(b.AsBoolean()),
            _ => 0,
        };
    }

    /// <summary>A legacy lookup value: an empty cell or an omitted argument is 0.</summary>
    public static CellValue LegacyTarget(CellValue target) =>
        target.Kind is CellValueKind.Empty or CellValueKind.Missing ? CellValue.Number(0) : target;

    /// <summary>
    /// MATCH type 0, VLOOKUP and HLOOKUP with FALSE: the first value of the same type that is
    /// equal; text with <c>*</c>, <c>?</c> or <c>~</c> is a pattern. Returns -1 when nothing matches.
    /// </summary>
    public static int FindExact(LookupVector vector, CellValue target, EvaluationContext context)
    {
        var culture = context.Culture;
        var pattern = target.Kind == CellValueKind.Text && target.AsText().AsSpan().IndexOfAny("*?~") >= 0 ? target.AsText() : null;
        var rank = Rank(target);
        var visited = 0;
        foreach (var (index, value) in vector.Occupied(0, vector.Length - 1, context))
        {
            if (++visited % CancellationCheckInterval == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            if (Rank(value) != rank)
                continue;
            if (pattern is not null ? Wildcard.Matches(pattern, value.AsText(), culture) : Compare(value, target, culture) == 0)
                return index;
        }

        return -1;
    }

    /// <summary>
    /// MATCH type 1 or -1, VLOOKUP and HLOOKUP with TRUE, LOOKUP: Excel's binary search over values
    /// of the target's type. Ascending data gives the last value not greater than the target,
    /// descending data the last value not less than it. Returns -1 when there is none.
    /// </summary>
    public static int FindSorted(LookupVector vector, CellValue target, bool descending, EvaluationContext context)
    {
        var rank = Rank(target);
        return BinarySearch(vector, target, descending ? order => order >= 0 : order => order <= 0, value => Rank(value) == rank, context);
    }

    /// <summary>
    /// XLOOKUP and XMATCH. Returns -1 when nothing matches, or null (<c>#VALUE!</c>) for an invalid
    /// regular expression or a pattern with binary search.
    /// </summary>
    public static int? Find(LookupVector vector, CellValue target, int matchMode, int searchMode, EvaluationContext context)
    {
        if (target.Kind is CellValueKind.Empty or CellValueKind.Missing)
            return FindEmpty(vector, searchMode != LastToFirst, context);

        var culture = context.Culture;
        if (matchMode is WildcardMatch or RegexMatch)
        {
            // A pattern cannot be searched for in sorted order.
            if (searchMode is BinaryAscending or BinaryDescending)
                return null;
            if (target.Kind != CellValueKind.Text)
                return Linear(vector, target, ExactMatch, searchMode != LastToFirst, context);
            Func<string, bool> matches;
            if (matchMode == WildcardMatch)
            {
                var pattern = target.AsText();
                matches = text => Wildcard.Matches(pattern, text, culture);
            }
            else
            {
                Regex regex;
                try
                {
                    regex = new Regex(target.AsText(), RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
                }
                catch (ArgumentException)
                {
                    return null;
                }

                matches = regex.IsMatch;
            }

            foreach (var (index, value) in InOrder(vector, searchMode != LastToFirst, context))
            {
                if (value.Kind == CellValueKind.Text && matches(value.AsText()))
                    return index;
            }

            return -1;
        }

        if (searchMode is FirstToLast or LastToFirst)
            return Linear(vector, target, matchMode, searchMode == FirstToLast, context);

        // Of equal values, ascending data gives the first and descending data the last: the search
        // finds the last value before the target's (below it, or for descending data not below
        // it), and the next comparable value is the far side.
        var descending = searchMode == BinaryDescending;
        Func<CellValue, bool> comparable = value => Rank(value) >= 0;
        var near = BinarySearch(vector, target, descending ? order => order >= 0 : order => order < 0, comparable, context);
        var far = NextComparable(vector, near + 1, vector.Length - 1, comparable, context, out var farValue);
        if (descending)
        {
            if (near >= 0 && Compare(vector.Get(near, context), target, culture) == 0)
                return near;
            return matchMode switch
            {
                ExactOrLarger => near,
                ExactOrSmaller => far,
                _ => -1,
            };
        }

        if (far >= 0 && Compare(farValue, target, culture) == 0)
            return far;
        return matchMode switch
        {
            ExactOrSmaller => near,
            ExactOrLarger => far,
            _ => -1,
        };
    }

    // Exact (equal value, same type) or the closest value on one side, by XLOOKUP's type order.
    // Ties keep the first one met in the search direction.
    private static int Linear(LookupVector vector, CellValue target, int matchMode, bool forward, EvaluationContext context)
    {
        var culture = context.Culture;
        var best = -1;
        var bestValue = default(CellValue);
        foreach (var (index, value) in InOrder(vector, forward, context))
        {
            if (Rank(value) < 0)
                continue;
            var order = Compare(value, target, culture);
            if (order == 0)
                return index;
            if (matchMode == ExactMatch || (matchMode == ExactOrSmaller ? order > 0 : order < 0))
                continue;
            if (best < 0 || (matchMode == ExactOrSmaller ? Compare(value, bestValue, culture) > 0 : Compare(value, bestValue, culture) < 0))
            {
                best = index;
                bestValue = value;
            }
        }

        return best;
    }

    // An empty lookup value finds the first (or last) empty position.
    private static int FindEmpty(LookupVector vector, bool forward, EvaluationContext context)
    {
        var expected = forward ? 0 : vector.Length - 1;
        foreach (var (index, value) in InOrder(vector, forward, context))
        {
            if (index != expected || value.Kind == CellValueKind.Empty)
                return expected;
            expected += forward ? 1 : -1;
        }

        return expected >= 0 && expected < vector.Length ? expected : -1;
    }

    private static IEnumerable<(int Index, CellValue Value)> InOrder(LookupVector vector, bool forward, EvaluationContext context)
    {
        var visited = 0;
        if (forward)
        {
            foreach (var item in vector.Occupied(0, vector.Length - 1, context))
            {
                if (++visited % CancellationCheckInterval == 0)
                    context.CancellationToken.ThrowIfCancellationRequested();
                yield return item;
            }

            yield break;
        }

        var all = new List<(int, CellValue)>(vector.Occupied(0, vector.Length - 1, context));
        for (var i = all.Count - 1; i >= 0; i--)
        {
            if (++visited % CancellationCheckInterval == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            yield return all[i];
        }
    }

    // Binary search that skips values it cannot compare: from the middle it moves right to the
    // first comparable value before the upper bound; if there is none, the right half is dropped.
    // Returns the last comparable position whose comparison with the target (-1, 0 or 1) satisfies
    // nearSide, or -1: the data is taken to be sorted so that nearSide holds up to some point.
    private static int BinarySearch(LookupVector vector, CellValue target, Func<int, bool> nearSide, Func<CellValue, bool> comparable, EvaluationContext context)
    {
        var culture = context.Culture;
        int low = 0, high = vector.Length - 1, best = -1, steps = 0;
        while (low <= high)
        {
            if (++steps % CancellationCheckInterval == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            var middle = low + ((high - low) / 2);
            var position = NextComparable(vector, middle, high, comparable, context, out var value);
            if (position < 0)
            {
                high = middle - 1;
                continue;
            }

            if (nearSide(Compare(value, target, culture)))
            {
                best = position;
                low = position + 1;
            }
            else
            {
                high = middle - 1;
            }
        }

        return best;
    }

    private static int NextComparable(LookupVector vector, int from, int to, Func<CellValue, bool> comparable, EvaluationContext context, out CellValue value)
    {
        var visited = 0;
        foreach (var (index, candidate) in vector.Occupied(from, to, context))
        {
            if (++visited % CancellationCheckInterval == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            if (comparable(candidate))
            {
                value = candidate;
                return index;
            }
        }

        value = default;
        return -1;
    }
}
