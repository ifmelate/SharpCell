using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using SharpCell.Parsing;

namespace SharpCell.Evaluation;

/// <summary>
/// Dirty tracking and recalculation for one workbook.
/// <para>
/// Editing a cell marks the formulas that read it, transitively, as dirty. Recalculation computes
/// dirty formulas on an explicit stack: an evaluation that reads a dirty cell collects every such
/// cell instead of recursing, they are computed first, and the formula is evaluated again. Chains
/// of any length therefore use no thread stack. A dirty cell met again while it waits for its
/// inputs closes a loop: every cell of the loop gets 0 and a diagnostic, as in Excel without
/// iterative calculation.
/// </para>
/// </summary>
internal sealed class Calculation(Workbook workbook)
{
    // Dirty formulas in the order they were marked; the IsDirty flag on the cell removes duplicates.
    private readonly List<CellKey> _dirty = [];
    private readonly HashSet<CellKey> _volatile = [];
    private readonly Dictionary<Worksheet, RangeIndex> _spillWatch = [];

    // Areas of array formulas (Ctrl+Shift+Enter): no cell inside may be changed on its own.
    private readonly Dictionary<Worksheet, RangeIndex> _fixedArrays = [];

    // Cells covered by spill and array formula areas, anchors included, against Workbook.MaxSpillCells.
    private long _spillCells;

    // Why the result being committed did not spill; reported with the cell's diagnostics.
    private string? _spillRefusal;

    // A cell evaluated this often within one calculation is not settling (spills that keep waking
    // each other): it is treated as a loop. Ordinary restarts stay far below.
    private const int MaxEvaluationsPerCell = 1000;
    private readonly Dictionary<CellKey, int> _evaluations = [];

    // Diagnostics describe the current state: a cell's entries go away when it is edited or
    // calculated without problems. Entries from Evaluate last until the next Evaluate.
    private readonly Dictionary<CellKey, IReadOnlyList<CalculationDiagnostic>> _cellDiagnostics = [];
    private IReadOnlyList<CalculationDiagnostic> _detachedDiagnostics = [];

    public DependencyGraph Graph { get; } = new();

    /// <summary>Formula evaluations attempted by recalculation, restarts included. For tests.</summary>
    public int EvaluationCount { get; private set; }

    public IReadOnlyList<CalculationDiagnostic> Diagnostics
    {
        get
        {
            var seen = new HashSet<CalculationDiagnostic>(ReferenceEqualityComparer.Instance);
            var result = new List<CalculationDiagnostic>();
            foreach (var entries in _cellDiagnostics.Values)
            {
                foreach (var diagnostic in entries)
                {
                    if (seen.Add(diagnostic))
                        result.Add(diagnostic);
                }
            }

            result.AddRange(_detachedDiagnostics);
            return result;
        }
    }

    /// <summary>Call before a cell's content changes: its old dependencies and its spill stop counting.</summary>
    public void BeforeChange(CellKey key, CellData? data)
    {
        if (data?.Registered is { } registered)
        {
            Graph.Unregister(key, registered);
            data.Registered = null;
        }

        if (data is not null)
            Unwatch(key, data);
        if (data?.SpillArea is { } spill)
        {
            ClearSpill(key, data);
            InvalidateArea(key.Sheet, spill, key, quiet: new HashSet<CellKey> { key });
        }

        if (data?.FixedArray is { } fixedArea && _fixedArrays.TryGetValue(key.Sheet, out var arrays))
            arrays.Remove(fixedArea, key);

        _volatile.Remove(key);
        _cellDiagnostics.Remove(key);
    }

    /// <summary>Call after a cell's content changed: it (if a formula) and everything reading it become dirty.</summary>
    public void AfterChange(CellKey key, CellData? data)
    {
        if (data?.Formula is not null)
            MarkDirty(key, data);
        Invalidate(key);

        // Content appeared or disappeared where an array result wants to spill.
        if (_spillWatch.TryGetValue(key.Sheet, out var watch))
        {
            var anchors = new List<CellKey>();
            watch.Query(key.Row, key.Column, anchors.Add);
            foreach (var anchor in anchors)
            {
                if (anchor != key && anchor.Data is { Formula: not null, IsDirty: false } anchorData)
                {
                    MarkDirty(anchor, anchorData);
                    Invalidate(anchor);
                }
            }
        }
    }

    /// <summary>Settings changed or a sheet appeared: every formula is dirty.</summary>
    public void InvalidateAll()
    {
        foreach (var sheet in workbook.Sheets)
        {
            foreach (var cell in sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
            {
                if (cell.Data.Formula is not null)
                    MarkDirty(new CellKey(sheet, cell.Row, cell.Column), cell.Data);
            }
        }
    }

    public void InvalidateName(string upperName)
    {
        foreach (var user in Graph.UsersOf(upperName))
        {
            if (user.Data is { Formula: not null } data)
            {
                MarkDirty(user, data);
                Invalidate(user);
            }
        }
    }

    /// <summary>Marks the formulas inside an area out of date, with everything that reads them.</summary>
    public void InvalidateArea(Worksheet sheet, Area area)
    {
        foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
        {
            if (cell.Data.Formula is null)
                continue;
            var key = new CellKey(sheet, cell.Row, cell.Column);
            MarkDirty(key, cell.Data);
            Invalidate(key);
        }
    }

    public void Recalculate(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        _evaluations.Clear();

        foreach (var key in (CellKey[])[.. _volatile])
        {
            if (key.Data is { Formula: not null } data)
            {
                MarkDirty(key, data);
                Invalidate(key);
            }
        }

        for (var i = 0; i < _dirty.Count; i++)
            Compute(_dirty[i], cancellationToken);
        _dirty.Clear();
    }

    /// <summary>Evaluates a formula outside any cell, computing dirty cells it reads first.</summary>
    public CellValue EvaluateDetached(FormulaNode formula, Worksheet? sheet, CellAddress origin, CancellationToken cancellationToken)
    {
        _evaluations.Clear();
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var context = new EvaluationContext(workbook, sheet, origin) { IsDetached = true, CancellationToken = cancellationToken };
            var value = EvaluateGuarded(formula, context);
            if (context.Pending.Count == 0)
            {
                _detachedDiagnostics = [.. context.Diagnostics];
                return value;
            }

            foreach (var key in context.Pending)
                Compute(key, cancellationToken);
        }
    }

    /// <summary>A formula was put in place by a file loader: it is calculated by the next recalculation.</summary>
    public void MarkLoaded(CellKey key, CellData data) => MarkDirty(key, data);

    /// <summary>Claims room in the spill budget for an area a file says a formula covers.</summary>
    public bool TryReserveLoadedArea(Area area) => TryReserve(area);

    /// <summary>Records an array formula's area, so no cell inside it can be changed on its own.</summary>
    public void RegisterFixedArray(CellKey anchor, Area area)
    {
        if (!_fixedArrays.TryGetValue(anchor.Sheet, out var arrays))
            _fixedArrays[anchor.Sheet] = arrays = new RangeIndex();
        arrays.Add(area, anchor);
    }

    /// <summary>The array formula whose area holds the cell, other than at its own top-left cell; or null.</summary>
    public CellKey? FixedArrayAt(Worksheet sheet, int row, int column)
    {
        if (!_fixedArrays.TryGetValue(sheet, out var arrays))
            return null;
        CellKey? found = null;
        arrays.Query(row, column, anchor =>
        {
            if (anchor.Row != row || anchor.Column != column)
                found = anchor;
        });
        return found;
    }

    private bool TryReserve(Area area)
    {
        if (area.CellCount > workbook.MaxSpillCells - _spillCells)
            return false;
        _spillCells += area.CellCount;
        return true;
    }

    private void MarkDirty(CellKey key, CellData data)
    {
        if (data.IsDirty)
            return;
        data.IsDirty = true;
        _dirty.Add(key);
    }

    // Breadth-first over readers: everything that (transitively) reads the cell becomes dirty. An
    // anchor's spilled cells change with it, so readers of its spill area are readers of the anchor.
    private void Invalidate(CellKey start)
    {
        var queue = new Queue<CellKey>();
        queue.Enqueue(start);
        void Visit(CellKey reader)
        {
            if (reader.Data is { Formula: not null, IsDirty: false } data)
            {
                MarkDirty(reader, data);
                queue.Enqueue(reader);
            }
        }

        while (queue.Count > 0)
        {
            var cell = queue.Dequeue();
            Graph.ForEachReader(cell, Visit);
            if (cell.Data?.SpillArea is { } spill)
                Graph.ForEachReader(cell.Sheet, spill, Visit);
        }
    }

    private void Compute(CellKey root, CancellationToken cancellationToken)
    {
        var stack = new List<(CellKey Key, CellKey? Requester)> { (root, null) };
        try
        {
            while (stack.Count > 0)
            {
                var (key, requester) = stack[^1];
                var data = key.Data;
                if (data?.Formula is null || !data.IsDirty)
                {
                    stack.RemoveAt(stack.Count - 1);
                    continue;
                }

                cancellationToken.ThrowIfCancellationRequested();
                var evaluations = _evaluations.TryGetValue(key, out var count) ? count + 1 : 1;
                _evaluations[key] = evaluations;
                if (evaluations > MaxEvaluationsPerCell)
                {
                    ResolveNonConvergence(key, data);
                    stack.RemoveAt(stack.Count - 1);
                    continue;
                }

                if (!data.InProgress)
                {
                    data.InProgress = true;
                    data.Requester = requester;
                }

                var context = new EvaluationContext(workbook, key.Sheet, key.Address)
                {
                    CancellationToken = cancellationToken,
                    Dependencies = new Dependencies(),
                    Legacy = data.IsLegacy,
                    LegacyFormula = data.IsLegacy,
                };
                EvaluationCount++;
                var value = EvaluateGuarded(data.Formula, context);
                data.LastAttempt = context.Dependencies;
                data.LastAttemptVolatile = context.UsedVolatile;

                if (context.Pending.Count == 0)
                {
                    CommitGuarded(key, data, value, context, cancellationToken);
                    stack.RemoveAt(stack.Count - 1);
                    continue;
                }

                foreach (var input in context.Pending)
                {
                    if (input.Data is { InProgress: true })
                    {
                        ResolveCycle(key, input);
                        break;
                    }

                    stack.Add((input, key));
                }
            }
        }
        finally
        {
            // After cancellation or a bug, nothing may stay marked as in progress.
            foreach (var (key, _) in stack)
            {
                if (key.Data is { } data)
                {
                    data.InProgress = false;
                    data.Requester = null;
                }
            }
        }
    }

    // An array result fills the rectangle below and right of its anchor, unless the rectangle leaves
    // the sheet or holds anything else (#SPILL!, neighbours untouched). The anchor watches its
    // rectangle, so clearing a blocking cell makes it try again. Returns the anchor's own value.
    private CellValue Spill(CellKey key, CellData data, CellValue[,] array, CancellationToken cancellationToken)
    {
        var rows = array.GetLength(0);
        var columns = array.GetLength(1);
        var lastRow = (long)key.Row + rows - 1;
        var lastColumn = (long)key.Column + columns - 1;

        // Excel does not spill inside a table: a formula there that returns several values is #SPILL!.
        if ((rows > 1 || columns > 1) && workbook.TableAt(key.Sheet, key.Row, key.Column) is not null)
            return CellValue.Error(ErrorKind.Spill);

        if (lastRow > CellAddress.MaxRow || lastColumn > CellAddress.MaxColumn)
            return CellValue.Error(ErrorKind.Spill);

        var area = new Area(key.Row, key.Column, (int)lastRow, (int)lastColumn);
        Watch(key, data, area);
        var store = key.Sheet.Store;
        foreach (var cell in store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
        {
            if (cell.Row != key.Row || cell.Column != key.Column)
                return CellValue.Error(ErrorKind.Spill);
        }

        if (!TryReserve(area))
        {
            _spillRefusal = $"The result has {area.CellCount} cells; spills of this workbook may cover {workbook.MaxSpillCells} cells in all.";
            return CellValue.Error(ErrorKind.Spill);
        }

        // Set before writing, so an interrupted write can be cleaned up by ClearSpill.
        data.SpillArea = area;
        var written = 0;
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                if (r == 0 && c == 0)
                    continue;
                if (++written % 4096 == 0)
                    cancellationToken.ThrowIfCancellationRequested();
                var spilled = store.GetOrCreate(key.Row + r, key.Column + c);
                spilled.Value = SpilledValue(array[r, c]);
                spilled.SpillAnchor = key;
            }
        }

        return SpilledValue(array[0, 0]);
    }

    // An array formula fills exactly its area: a single value or a single row or column repeats,
    // cells beyond a larger result are #N/A, and a result larger than the area is cut. The area
    // is the formula's own, so nothing can block it.
    private CellValue FillFixed(CellKey key, CellData data, CellValue value, Area area, CancellationToken cancellationToken)
    {
        if (!TryReserve(area))
        {
            _spillRefusal = $"The array formula covers {area.CellCount} cells; spills of this workbook may cover {workbook.MaxSpillCells} cells in all.";
            return CellValue.Error(ErrorKind.Spill);
        }

        CellValue[,]? array = value.Kind == CellValueKind.Array ? value.AsArray() : null;
        CellValue Element(int r, int c)
        {
            if (array is null)
                return value;
            var rows = array.GetLength(0);
            var columns = array.GetLength(1);
            var row = rows == 1 ? 0 : r;
            var column = columns == 1 ? 0 : c;
            return row < rows && column < columns ? array[row, column] : CellValue.Error(ErrorKind.NA);
        }

        var store = key.Sheet.Store;
        data.SpillArea = area;
        var written = 0;
        for (var r = 0; r < area.Rows; r++)
        {
            for (var c = 0; c < area.Columns; c++)
            {
                if (r == 0 && c == 0)
                    continue;
                if (++written % 4096 == 0)
                    cancellationToken.ThrowIfCancellationRequested();

                // A formula inside the area can only come from a malformed file; it is left alone.
                if (store.Get(area.FirstRow + r, area.FirstColumn + c) is { Formula: not null })
                    continue;
                var member = store.GetOrCreate(area.FirstRow + r, area.FirstColumn + c);
                member.Value = SpilledValue(Element(r, c));
                member.SpillAnchor = key;
            }
        }

        return SpilledValue(Element(0, 0));
    }

    private static CellValue SpilledValue(CellValue value) => value.Kind switch
    {
        CellValueKind.Empty or CellValueKind.Missing => CellValue.Number(0),
        CellValueKind.Lambda or CellValueKind.Array => CellValue.Error(ErrorKind.Calc),
        _ => value,
    };

    private void ClearSpill(CellKey key, CellData data)
    {
        if (data.SpillArea is not { } area)
            return;

        var store = key.Sheet.Store;
        var owned = new List<StoredCell>();
        foreach (var cell in store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
        {
            if (cell.Data.SpillAnchor == key)
                owned.Add(cell);
        }

        // Right to left: a row keeps its columns in a list, and removing from its end is cheap.
        for (var i = owned.Count - 1; i >= 0; i--)
            store.Remove(owned[i].Row, owned[i].Column);
        data.SpillArea = null;
        _spillCells -= area.CellCount;
    }

    // Only stored cells: an area read from a file can claim far more cells than exist.
    private static List<(int Row, int Column, CellValue Value)> Snapshot(Worksheet sheet, Area area)
    {
        var values = new List<(int, int, CellValue)>();
        foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
            values.Add((cell.Row, cell.Column, cell.Data.Value));
        return values;
    }

    private static bool SameValues(List<(int Row, int Column, CellValue Value)> a, List<(int Row, int Column, CellValue Value)> b)
    {
        if (a.Count != b.Count)
            return false;
        for (var i = 0; i < a.Count; i++)
        {
            if (a[i].Row != b[i].Row || a[i].Column != b[i].Column || !a[i].Value.Equals(b[i].Value))
                return false;
        }

        return true;
    }

    private void Watch(CellKey key, CellData data, Area area)
    {
        if (!_spillWatch.TryGetValue(key.Sheet, out var watch))
            _spillWatch[key.Sheet] = watch = new RangeIndex();
        watch.Add(area, key);
        data.SpillWatch = area;
    }

    private void Unwatch(CellKey key, CellData data)
    {
        if (data.SpillWatch is { } area && _spillWatch.TryGetValue(key.Sheet, out var watch))
            watch.Remove(area, key);
        data.SpillWatch = null;
    }

    // The spill of `anchor` left or covered `area`: its readers (the anchor too, if it reads its own
    // spill) and other anchors wanting to spill there must be calculated again.
    private void InvalidateArea(Worksheet sheet, Area area, CellKey anchor, IReadOnlySet<CellKey>? quiet)
    {
        var affected = new List<CellKey>();
        Graph.ForEachReader(sheet, area, affected.Add);
        if (_spillWatch.TryGetValue(sheet, out var watch))
            watch.QueryOverlap(area, other => { if (other != anchor) affected.Add(other); });

        foreach (var cell in affected)
        {
            if (quiet is not null && quiet.Contains(cell))
                continue;
            if (cell.Data is { Formula: not null, IsDirty: false } data)
            {
                MarkDirty(cell, data);
                Invalidate(cell);
            }
        }
    }

    // The last line of defence: whatever goes wrong inside one formula (a bug, an allocation that
    // fails) turns into #VALUE! for that cell, so one cell cannot stop the workbook from calculating.
    private static CellValue EvaluateGuarded(FormulaNode formula, EvaluationContext context)
    {
        try
        {
            return Evaluator.EvaluateFormula(formula, context);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            context.Report(DiagnosticKind.FunctionFailure, $"Evaluation failed: {ex.GetType().Name}: {ex.Message}");
            return CellValue.Error(ErrorKind.Value);
        }
    }

    private void SetDiagnostics(CellKey key, IReadOnlyList<CalculationDiagnostic> diagnostics)
    {
        if (diagnostics.Count == 0)
            _cellDiagnostics.Remove(key);
        else
            _cellDiagnostics[key] = [.. diagnostics];
    }

    // Committing writes spilled cells, which can fail on huge arrays: the cell then gets #VALUE!
    // with nothing half-written left behind. Cancellation leaves it dirty and clean of spill.
    private void CommitGuarded(CellKey key, CellData data, CellValue value, EvaluationContext context, CancellationToken cancellationToken)
    {
        try
        {
            _spillRefusal = null;
            Commit(key, data, value, context.Dependencies!, context.UsedVolatile, cancellationToken);
            SetDiagnostics(key, _spillRefusal is { } refusal
                ? [.. context.Diagnostics, new CalculationDiagnostic(DiagnosticKind.LimitExceeded, key.Sheet, key.Address.ToString(), refusal)]
                : context.Diagnostics);
        }
        catch (OperationCanceledException)
        {
            ClearSpill(key, data);
            throw;
        }
        catch (Exception ex)
        {
            ClearSpill(key, data);
            Commit(key, data, CellValue.Error(ErrorKind.Value), context.Dependencies!, context.UsedVolatile);
            SetDiagnostics(key, [new CalculationDiagnostic(DiagnosticKind.FunctionFailure, key.Sheet, key.Address.ToString(),
                $"Storing the result failed: {ex.GetType().Name}: {ex.Message}")]);
        }
    }

    /// <param name="quiet">Cells not to mark dirty again; the members of a loop being resolved.</param>
    /// <param name="key">The cell.</param>
    /// <param name="data">The cell's stored formula and state.</param>
    /// <param name="value">The value the formula produced.</param>
    /// <param name="dependencies">The cells and ranges the formula read, registered in the dependency graph.</param>
    /// <param name="usedVolatile">Whether the formula called a volatile function, so it recalculates every time.</param>
    /// <param name="cancellationToken">Cancels spilling a large result.</param>
    private void Commit(CellKey key, CellData data, CellValue value, Dependencies dependencies, bool usedVolatile,
        CancellationToken cancellationToken = default, IReadOnlySet<CellKey>? quiet = null)
    {
        var oldSpill = data.SpillArea;
        var oldValues = oldSpill is { } previousArea ? Snapshot(key.Sheet, previousArea) : null;
        Unwatch(key, data);
        if (oldSpill is not null)
            ClearSpill(key, data);
        if (data.FixedArray is { } fixedArea)
            value = FillFixed(key, data, value, fixedArea, cancellationToken);
        else if (value.Kind == CellValueKind.Array)
            value = Spill(key, data, value.AsArray(), cancellationToken);

        data.Value = value;
        data.IsDirty = false;
        data.InProgress = false;
        data.Requester = null;
        if (data.Registered is { } previous)
            Graph.Unregister(key, previous);
        data.Registered = dependencies;
        Graph.Register(key, dependencies);
        if (usedVolatile)
            _volatile.Add(key);
        else
            _volatile.Remove(key);

        // Cells the spill left or newly covers changed for whoever reads them. This runs after the
        // anchor is clean, so an anchor that reads its own spill becomes dirty again and the loop
        // is detected on the next evaluation instead of leaving a stale value.
        // An identical spill (same area, same values) changes nothing for anyone.
        if (oldSpill is { } same && data.SpillArea == same && SameValues(oldValues!, Snapshot(key.Sheet, same)))
            return;
        if (oldSpill is { } before)
            InvalidateArea(key.Sheet, before, key, quiet);
        if (data.SpillArea is { } after && after != oldSpill)
            InvalidateArea(key.Sheet, after, key, quiet);
    }

    private void ResolveNonConvergence(CellKey key, CellData data)
    {
        Commit(key, data, CellValue.Number(0), data.LastAttempt ?? new Dependencies(), data.LastAttemptVolatile,
            quiet: new HashSet<CellKey> { key });
        SetDiagnostics(key, [new CalculationDiagnostic(DiagnosticKind.CircularReference, key.Sheet, key.Address.ToString(),
            $"Circular reference: {key} does not settle (spills keep invalidating each other)")]);
    }

    // The loop is the chain of requesters from the cell that asked back up to the cell it asked for.
    private void ResolveCycle(CellKey asker, CellKey asked)
    {
        var members = new List<CellKey>();
        CellKey? current = asker;
        while (current is { } member)
        {
            members.Add(member);
            if (member == asked)
                break;
            current = member.Data?.Requester;
        }

        var path = new StringBuilder("Circular reference: ");
        for (var i = members.Count - 1; i >= 0; i--)
            path.Append(members[i]).Append(" -> ");
        path.Append(members[^1]);
        var diagnostic = new CalculationDiagnostic(DiagnosticKind.CircularReference, asked.Sheet, asked.Address.ToString(), path.ToString());

        // Resolving the loop must not wake its own members again (they would spill, read each
        // other's spill and loop forever); everything else that depended on them is invalidated.
        var quiet = new HashSet<CellKey>(members);
        foreach (var member in members)
        {
            var data = member.Data!;
            Commit(member, data, CellValue.Number(0), data.LastAttempt ?? new Dependencies(), data.LastAttemptVolatile, quiet: quiet);
            SetDiagnostics(member, [diagnostic]);
        }
    }
}
