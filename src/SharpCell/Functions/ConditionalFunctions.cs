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

    /// <summary>Positions (row, column offsets) where the grid may hold something other than empty.</summary>
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
    /// position where all criteria hold. Positions where every range is empty are not visited one
    /// by one: their number is passed to <paramref name="emptyMatches"/> if empty cells meet all
    /// criteria. Ranges of different shapes are <c>#VALUE!</c>.
    /// </summary>
    /// <returns>Null, or the error to return.</returns>
    public static CellValue? Visit(FunctionCall call, IReadOnlyList<(ValueGrid Range, Criterion Criterion)> conditions,
        ValueGrid? target, Action<CellValue> match, Action<long> emptyMatches)
    {
        var context = call.Context;
        var rows = conditions[0].Range.Rows;
        var columns = conditions[0].Range.Columns;
        foreach (var (range, _) in conditions)
        {
            if (range.Rows != rows || range.Columns != columns)
                return CellValue.Error(ErrorKind.Value);
        }

        if (target is { } t && (t.Rows != rows || t.Columns != columns))
            return CellValue.Error(ErrorKind.Value);

        var occupied = new HashSet<(int, int)>();
        foreach (var (range, _) in conditions)
            range.AddOccupied(occupied);
        target?.AddOccupied(occupied);

        // Row by row, so sums add up in the order Excel adds them.
        var positions = new List<(int Row, int Column)>(occupied);
        positions.Sort();
        var visited = 0;
        var emptyCandidates = (long)rows * columns - positions.Count;
        foreach (var (row, column) in positions)
        {
            if (++visited % CancellationCheckInterval == 0)
                context.CancellationToken.ThrowIfCancellationRequested();

            var all = true;
            foreach (var (range, criterion) in conditions)
            {
                if (!criterion.Matches(range.Get(row, column, context)))
                {
                    all = false;
                    break;
                }
            }

            if (all)
                match(target is { } values ? values.Get(row, column, context) : CellValue.Empty);
        }

        if (emptyCandidates > 0)
        {
            var allEmpty = true;
            foreach (var (_, criterion) in conditions)
                allEmpty &= criterion.MatchesEmpty;
            if (allEmpty)
                emptyMatches(emptyCandidates);
        }

        return null;
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
