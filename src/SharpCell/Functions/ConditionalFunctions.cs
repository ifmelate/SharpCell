using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// A range or array argument seen as a grid of values with a fixed shape. A reference reads its
/// cells through the evaluation context (dirty cells, dependencies); only stored cells exist.
/// </summary>
internal readonly struct ValueGrid
{
    private readonly Worksheet? _sheet;
    private readonly Area _area;
    private readonly CellValue[,]? _array;
    private readonly CellValue _scalar;

    private ValueGrid(Worksheet? sheet, Area area, CellValue[,]? array, CellValue scalar)
    {
        _sheet = sheet;
        _area = area;
        _array = array;
        _scalar = scalar;
        Rows = array?.GetLength(0) ?? area.Rows;
        Columns = array?.GetLength(1) ?? area.Columns;
    }

    public int Rows { get; }

    public int Columns { get; }

    public bool IsReference => _sheet is not null;

    /// <summary>A single-area reference, an array or a scalar; anything else is <c>#VALUE!</c>.</summary>
    public static bool TryCreate(Operand operand, out ValueGrid grid, out CellValue error)
    {
        error = default;
        grid = default;
        if (operand.Reference is { } reference)
        {
            if (!reference.IsSingleArea)
            {
                error = CellValue.Error(ErrorKind.Value);
                return false;
            }

            var (sheet, area) = reference.Areas[0];
            grid = new ValueGrid(sheet, area, null, default);
            return true;
        }

        var value = operand.Value;
        if (value.IsError)
        {
            error = value;
            return false;
        }

        grid = value.Kind == CellValueKind.Array
            ? new ValueGrid(null, default, value.AsArray(), default)
            : new ValueGrid(null, new Area(1, 1, 1, 1), null, value);
        return true;
    }

    /// <summary>
    /// The same top-left cell with another shape, as SUMIF reads its sum range: SUMIF(A1:A9,"x",B1)
    /// sums B1:B9. The new area must be recorded as a dependency. Null if it leaves the sheet.
    /// </summary>
    public ValueGrid? Resized(int rows, int columns, EvaluationContext context)
    {
        if (_sheet is null)
            return rows <= Rows && columns <= Columns ? this : null;

        var lastRow = (long)_area.FirstRow + rows - 1;
        var lastColumn = (long)_area.FirstColumn + columns - 1;
        if (lastRow > CellAddress.MaxRow || lastColumn > CellAddress.MaxColumn)
            return null;

        var area = new Area(_area.FirstRow, _area.FirstColumn, (int)lastRow, (int)lastColumn);
        context.RecordReference(new Reference(_sheet, area));
        return new ValueGrid(_sheet, area, null, default);
    }

    public CellValue Get(int row, int column, EvaluationContext context)
    {
        if (_array is not null)
            return _array[row, column];
        if (_sheet is null)
            return _scalar;
        return context.ReadCell(_sheet, _area.FirstRow + row, _area.FirstColumn + column);
    }

    /// <summary>
    /// Positions (row, column offsets) where the grid may hold something other than empty; used by
    /// the database, distribution test and array math functions.
    /// </summary>
    public void AddOccupied(HashSet<(int, int)> positions)
    {
        if (_sheet is null)
        {
            for (var r = 0; r < Rows; r++)
            {
                for (var c = 0; c < Columns; c++)
                    positions.Add((r, c));
            }

            return;
        }

        foreach (var cell in _sheet.Store.Enumerate(_area.FirstRow, _area.FirstColumn, _area.LastRow, _area.LastColumn))
            positions.Add((cell.Row - _area.FirstRow, cell.Column - _area.FirstColumn));
    }

    /// <summary>
    /// An array, a scalar, or a reference with no more positions than its sheet has cells, so
    /// walking every position costs no more than walking the sheet.
    /// </summary>
    public bool IsNoLargerThanItsSheet => _sheet is null || (long)Rows * Columns <= _sheet.Store.Count;

    /// <summary>
    /// The cells a reference holds, row by row, columns ascending; null for an array or a scalar,
    /// which hold a value at every position.
    /// </summary>
    public IEnumerator<StoredCell>? StoredCells() =>
        _sheet?.Store.Enumerate(_area.FirstRow, _area.FirstColumn, _area.LastRow, _area.LastColumn).GetEnumerator();

    /// <summary>Reads a cell of <see cref="StoredCells"/> through the evaluation context (dirty cells, dependencies).</summary>
    public CellValue Read(StoredCell cell, EvaluationContext context) => context.ReadCell(_sheet!, cell);

    /// <summary>The row and column offsets of a cell of <see cref="StoredCells"/>.</summary>
    public (int Row, int Column) OffsetOf(StoredCell cell) => (cell.Row - _area.FirstRow, cell.Column - _area.FirstColumn);
}

/// <summary>
/// The shared machinery of COUNTIF, SUMIF, the *IFS functions and similar: ranges paired with
/// criteria, all of one shape, and optionally a range of values to aggregate.
/// </summary>
internal static class Conditional
{
    private const int CancellationCheckInterval = 4096;

    /// <summary>
    /// Calls <paramref name="match"/> with the target value (or Empty without a target) at every
    /// position where all criteria hold, row by row, columns ascending. Positions where every range
    /// is empty are not visited one by one when empty cells meet all criteria: their number is
    /// passed to <paramref name="emptyMatches"/>. Ranges of different shapes are <c>#VALUE!</c>.
    /// </summary>
    /// <returns>Null, or the error to return.</returns>
    public static CellValue? Visit(FunctionCall call, IReadOnlyList<(ValueGrid Range, Criterion Criterion)> conditions,
        ValueGrid? target, Action<CellValue> match, Action<long> emptyMatches)
    {
        var context = call.Context;
        var count = conditions.Count;
        var rows = conditions[0].Range.Rows;
        var columns = conditions[0].Range.Columns;
        for (var i = 0; i < count; i++)
        {
            var range = conditions[i].Range;
            if (range.Rows != rows || range.Columns != columns)
                return CellValue.Error(ErrorKind.Value);
        }

        if (target is { } t && (t.Rows != rows || t.Columns != columns))
            return CellValue.Error(ErrorKind.Value);

        var criteria = new Criterion[count];
        var allMatchEmpty = true;
        var small = true;
        for (var i = 0; i < count; i++)
        {
            criteria[i] = conditions[i].Criterion;
            allMatchEmpty &= criteria[i].MatchesEmpty;
            small &= conditions[i].Range.IsNoLargerThanItsSheet;
        }

        // When some criterion fails on empty, a position where every range is empty never matches,
        // so visiting it changes nothing. Ranges no larger than their sheets are then cheaper to
        // walk position by position, reading a range only when the criteria before it held.
        if (!allMatchEmpty && small)
        {
            VisitEvery(context, conditions, criteria, target, match);
            return null;
        }

        var emptyCandidates = (long)rows * columns - VisitStored(context, conditions, criteria, target, match);
        if (emptyCandidates > 0 && allMatchEmpty)
            emptyMatches(emptyCandidates);
        return null;
    }

    private static void VisitEvery(EvaluationContext context, IReadOnlyList<(ValueGrid Range, Criterion Criterion)> conditions,
        Criterion[] criteria, ValueGrid? target, Action<CellValue> match)
    {
        var grids = new ValueGrid[criteria.Length];
        for (var i = 0; i < grids.Length; i++)
            grids[i] = conditions[i].Range;

        var rows = grids[0].Rows;
        var columns = grids[0].Columns;
        var visited = 0L;
        for (var row = 0; row < rows; row++)
        {
            for (var column = 0; column < columns; column++)
            {
                if (++visited % CancellationCheckInterval == 0)
                    context.CancellationToken.ThrowIfCancellationRequested();

                var all = true;
                for (var i = 0; i < grids.Length; i++)
                {
                    if (!criteria[i].Matches(grids[i].Get(row, column, context)))
                    {
                        all = false;
                        break;
                    }
                }

                if (all)
                    match(target is { } values ? values.Get(row, column, context) : CellValue.Empty);
            }
        }
    }

    // The positions where some range holds a cell, row by row; returns how many there were.
    private static long VisitStored(EvaluationContext context, IReadOnlyList<(ValueGrid Range, Criterion Criterion)> conditions,
        Criterion[] criteria, ValueGrid? target, Action<CellValue> match)
    {
        // One cursor per range, the target last. Each walks its own stored cells in order, so the
        // positions where any range holds a cell come out merged, row by row, as Excel adds them.
        var count = criteria.Length;
        var columns = conditions[0].Range.Columns;
        var cursors = new Cursor[target is null ? count : count + 1];
        long visited = 0;
        try
        {
            for (var i = 0; i < count; i++)
                cursors[i] = new Cursor(conditions[i].Range, columns);
            if (target is { } values)
                cursors[count] = new Cursor(values, columns);

            while (true)
            {
                var key = Cursor.End;
                int row = 0, column = 0;
                for (var i = 0; i < cursors.Length; i++)
                {
                    ref var cursor = ref cursors[i];
                    if (cursor.Key < key)
                    {
                        key = cursor.Key;
                        row = cursor.Row;
                        column = cursor.Column;
                    }
                }

                if (key == Cursor.End)
                    return visited;
                if (++visited % CancellationCheckInterval == 0)
                    context.CancellationToken.ThrowIfCancellationRequested();

                var all = true;
                for (var i = 0; i < count; i++)
                {
                    if (!criteria[i].Matches(cursors[i].ValueAt(key, row, column, context)))
                    {
                        all = false;
                        break;
                    }
                }

                if (all)
                    match(target is null ? CellValue.Empty : cursors[count].ValueAt(key, row, column, context));

                for (var i = 0; i < cursors.Length; i++)
                {
                    ref var cursor = ref cursors[i];
                    if (cursor.Key == key)
                        cursor.Advance();
                }
            }
        }
        finally
        {
            foreach (var cursor in cursors)
                cursor.Dispose();
        }
    }

    /// <summary>
    /// Walks the positions where a grid may hold something other than empty, row by row, columns
    /// ascending: the stored cells of a reference, every position of an array or a scalar.
    /// </summary>
    private struct Cursor
    {
        /// <summary>The key of a cursor that has passed its last position.</summary>
        public const long End = long.MaxValue;

        private readonly ValueGrid _grid;
        private readonly IEnumerator<StoredCell>? _cells;
        private readonly int _columns;
        private readonly long _positions;
        private StoredCell _cell;

        public Cursor(ValueGrid grid, int columns)
        {
            _grid = grid;
            _columns = columns;
            _positions = (long)grid.Rows * columns;
            _cells = grid.StoredCells();
            if (_cells is null)
                Key = _positions > 0 ? 0 : End;
            else
                Next();
        }

        /// <summary>The current position, row offset times the width plus column offset; <see cref="End"/> when done.</summary>
        public long Key { get; private set; }

        public int Row { get; private set; }

        public int Column { get; private set; }

        /// <summary>The value at <paramref name="key"/>: the current cell if the cursor is there, otherwise empty.</summary>
        public readonly CellValue ValueAt(long key, int row, int column, EvaluationContext context)
        {
            if (Key != key)
                return CellValue.Empty;
            return _cells is null ? _grid.Get(row, column, context) : _grid.Read(_cell, context);
        }

        public void Advance()
        {
            if (_cells is not null)
            {
                Next();
                return;
            }

            if (Key + 1 >= _positions)
            {
                Key = End;
                return;
            }

            Key++;
            if (++Column == _columns)
            {
                Column = 0;
                Row++;
            }
        }

        public readonly void Dispose() => _cells?.Dispose();

        private void Next()
        {
            if (!_cells!.MoveNext())
            {
                Key = End;
                return;
            }

            _cell = _cells.Current;
            var (row, column) = _grid.OffsetOf(_cell);
            Row = row;
            Column = column;
            Key = (long)row * _columns + column;
        }
    }

    /// <summary>Reads (range, criterion) pairs from the arguments starting at <paramref name="first"/>.</summary>
    public static bool TryReadPairs(FunctionCall call, int first, out List<(ValueGrid, Criterion)> conditions, out CellValue error)
    {
        conditions = [];
        error = default;
        if ((call.Count - first) % 2 != 0)
        {
            error = CellValue.Error(ErrorKind.Value);
            return false;
        }

        for (var i = first; i < call.Count; i += 2)
        {
            if (!ValueGrid.TryCreate(call[i], out var range, out error))
                return false;
            conditions.Add((range, Criterion.Parse(call.Value(i + 1), call.Context.Culture, call.Context.DateSystem)));
        }

        return true;
    }
}

/// <summary>COUNTIF, COUNTIFS, SUMIF, SUMIFS.</summary>
internal static class ConditionalFunctions
{
    // Ranges are taken as references; criteria are scalars, so an array of criteria gives an array of results.
    private static readonly ArgumentKind[] Pairs = BuildPairs(0);
    private static readonly ArgumentKind[] SumPairs = BuildPairs(1);

    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("COUNTIF", 2, 2, [ArgumentKind.Any, ArgumentKind.Value], CountIfs));
        registry.Add(new FunctionInfo("COUNTIFS", 2, FunctionRegistry.MaxArguments - 1, Pairs, CountIfs));
        registry.Add(new FunctionInfo("SUMIF", 2, 3, [ArgumentKind.Any, ArgumentKind.Value, ArgumentKind.Any], SumIf));
        registry.Add(new FunctionInfo("SUMIFS", 3, FunctionRegistry.MaxArguments, SumPairs, SumIfs));
    }

    private static ArgumentKind[] BuildPairs(int leading)
    {
        var kinds = new ArgumentKind[FunctionRegistry.MaxArguments];
        for (var i = 0; i < kinds.Length; i++)
            kinds[i] = i < leading || (i - leading) % 2 == 0 ? ArgumentKind.Any : ArgumentKind.Value;
        return kinds;
    }

    private static Operand CountIfs(FunctionCall call)
    {
        if (!Conditional.TryReadPairs(call, 0, out var conditions, out var error))
            return error;

        long count = 0;
        var failure = Conditional.Visit(call, conditions, null, _ => count++, empty => count += empty);
        return failure ?? CellValue.Number(count);
    }

    private static Operand SumIf(FunctionCall call)
    {
        if (!ValueGrid.TryCreate(call[0], out var range, out var error))
            return error;
        var criterion = Criterion.Parse(call.Value(1), call.Context.Culture, call.Context.DateSystem);
        var target = range;
        if (call.Has(2))
        {
            if (!ValueGrid.TryCreate(call[2], out var sum, out error))
                return error;
            if (sum.Resized(range.Rows, range.Columns, call.Context) is not { } resized)
                return CellValue.Error(ErrorKind.Value);
            target = resized;
        }

        return Sum(call, [(range, criterion)], target);
    }

    private static Operand SumIfs(FunctionCall call)
    {
        if (!ValueGrid.TryCreate(call[0], out var target, out var error))
            return error;
        if (!Conditional.TryReadPairs(call, 1, out var conditions, out error))
            return error;
        return Sum(call, conditions, target);
    }

    // Only numbers are added; an error in a matching cell is the result.
    private static Operand Sum(FunctionCall call, List<(ValueGrid, Criterion)> conditions, ValueGrid target)
    {
        var total = 0.0;
        CellValue? firstError = null;
        var failure = Conditional.Visit(call, conditions, target, value =>
        {
            if (value.Kind == CellValueKind.Number)
                total += value.AsNumber();
            else if (value.IsError)
                firstError ??= value;
        }, _ => { });
        return failure ?? firstError ?? CellValue.Number(total);
    }
}
