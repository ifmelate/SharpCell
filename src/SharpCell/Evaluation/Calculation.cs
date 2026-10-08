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

        if (data?.SpillArea is { } spill)
        {
            ClearSpill(key, data);
            InvalidateArea(key.Sheet, spill, except: key);
        }

        _volatile.Remove(key);
        _cellDiagnostics.Remove(key);
    }

    /// <summary>Call after a cell's content changed: it (if a formula) and everything reading it become dirty.</summary>
    public void AfterChange(CellKey key, CellData? data)
    {
        if (data?.Formula is not null)
            MarkDirty(key, data);
        Invalidate(key);
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

    public void Recalculate(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

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
    public CellValue EvaluateDetached(FormulaNode formula, Worksheet? sheet, CellAddress origin)
    {
        while (true)
        {
            var context = new EvaluationContext(workbook, sheet, origin) { IsDetached = true };
            var value = EvaluateGuarded(formula, context);
            if (context.Pending.Count == 0)
            {
                _detachedDiagnostics = [.. context.Diagnostics];
                return value;
            }

            foreach (var key in context.Pending)
                Compute(key, CancellationToken.None);
        }
    }

    private void MarkDirty(CellKey key, CellData data)
    {
        if (data.IsDirty)
            return;
        data.IsDirty = true;
        _dirty.Add(key);
    }

    // Breadth-first over readers: everything that (transitively) reads the cell becomes dirty.
    private void Invalidate(CellKey start)
    {
        var queue = new Queue<CellKey>();
        queue.Enqueue(start);
        while (queue.Count > 0)
        {
            Graph.ForEachReader(queue.Dequeue(), reader =>
            {
                if (reader.Data is { Formula: not null, IsDirty: false } data)
                {
                    MarkDirty(reader, data);
                    queue.Enqueue(reader);
                }
            });
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
                };
                EvaluationCount++;
                var value = EvaluateGuarded(data.Formula, context);
                data.LastAttempt = context.Dependencies;
                data.LastAttemptVolatile = context.UsedVolatile;

                if (context.Pending.Count == 0)
                {
                    Commit(key, data, value, context.Dependencies, context.UsedVolatile);
                    SetDiagnostics(key, context.Diagnostics);
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
    // the sheet or holds anything else (#SPILL!, neighbours untouched). The anchor reads its
    // rectangle, so clearing a blocking cell makes it try again. Returns the anchor's own value.
    private static CellValue Spill(CellKey key, CellData data, CellValue[,] array, Dependencies dependencies)
    {
        var rows = array.GetLength(0);
        var columns = array.GetLength(1);
        var lastRow = (long)key.Row + rows - 1;
        var lastColumn = (long)key.Column + columns - 1;
        if (lastRow > CellAddress.MaxRow || lastColumn > CellAddress.MaxColumn)
            return CellValue.Error(ErrorKind.Spill);

        var area = new Area(key.Row, key.Column, (int)lastRow, (int)lastColumn);
        dependencies.Areas.Add(new SheetArea(key.Sheet, area));
        var store = key.Sheet.Store;
        foreach (var cell in store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
        {
            if (cell.Row != key.Row || cell.Column != key.Column)
                return CellValue.Error(ErrorKind.Spill);
        }

        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                if (r == 0 && c == 0)
                    continue;
                var spilled = store.GetOrCreate(key.Row + r, key.Column + c);
                spilled.Value = SpilledValue(array[r, c]);
                spilled.SpillAnchor = key;
            }
        }

        data.SpillArea = area;
        return SpilledValue(array[0, 0]);
    }

    private static CellValue SpilledValue(CellValue value) => value.Kind switch
    {
        CellValueKind.Empty or CellValueKind.Missing => CellValue.Number(0),
        CellValueKind.Lambda or CellValueKind.Array => CellValue.Error(ErrorKind.Calc),
        _ => value,
    };

    private static void ClearSpill(CellKey key, CellData data)
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

        foreach (var cell in owned)
            store.Remove(cell.Row, cell.Column);
        data.SpillArea = null;
    }

    private void InvalidateArea(Worksheet sheet, Area area, CellKey except)
    {
        var readers = new List<CellKey>();
        Graph.ForEachReader(sheet, area, readers.Add);
        foreach (var reader in readers)
        {
            if (reader != except && reader.Data is { Formula: not null, IsDirty: false } data)
            {
                MarkDirty(reader, data);
                Invalidate(reader);
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

    private void Commit(CellKey key, CellData data, CellValue value, Dependencies dependencies, bool usedVolatile)
    {
        var oldSpill = data.SpillArea;
        if (oldSpill is not null)
            ClearSpill(key, data);
        if (value.Kind == CellValueKind.Array)
            value = Spill(key, data, value.AsArray(), dependencies);

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
        // anchor is clean, so a reader that the anchor itself depends on makes the anchor dirty
        // again and a loop through the spill is detected instead of leaving a stale value.
        if (oldSpill is { } before)
            InvalidateArea(key.Sheet, before, except: key);
        if (data.SpillArea is { } after && after != oldSpill)
            InvalidateArea(key.Sheet, after, except: key);
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

        foreach (var member in members)
        {
            var data = member.Data!;
            Commit(member, data, CellValue.Number(0), data.LastAttempt ?? new Dependencies(), data.LastAttemptVolatile);
            SetDiagnostics(member, [diagnostic]);
        }
    }
}
