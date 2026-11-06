using System;
using System.Collections.Generic;
using System.Linq;
using SharpCell.Parsing;

namespace SharpCell;

// Outlines, as Excel's Data > Group makes them: rows and columns at levels 1 to 7, each group with a
// summary row below it (or above) and a summary column right of it (or left), whose button collapses
// the group by hiding its rows or columns. Hidden rows count for SUBTOTAL as any hidden row does.
public sealed partial class Worksheet
{
    /// <summary>The deepest outline level Excel allows.</summary>
    public static int MaxOutlineLevel => 7;

    private readonly Dictionary<int, int> _rowLevels = [];
    private readonly Dictionary<int, int> _columnLevels = [];
    private readonly HashSet<int> _collapsedRows = [];
    private readonly HashSet<int> _collapsedColumns = [];
    private bool _summaryBelow = true;
    private bool _summaryRight = true;

    /// <summary>Whether a group's summary row is the one below it, as Excel has it unless Data &gt; Outline settings say otherwise.</summary>
    public bool OutlineSummaryBelow
    {
        get => _summaryBelow;
        set
        {
            ThrowIfCannotChange();
            _summaryBelow = value;
        }
    }

    /// <summary>Whether a group's summary column is the one right of it, as Excel has it unless set otherwise.</summary>
    public bool OutlineSummaryRight
    {
        get => _summaryRight;
        set
        {
            ThrowIfCannotChange();
            _summaryRight = value;
        }
    }

    /// <summary>A row's outline level: 0 outside any group, up to <see cref="MaxOutlineLevel"/>.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The row is outside the sheet.</exception>
    public int RowOutlineLevel(int row)
    {
        CheckRow(row);
        return _rowLevels.GetValueOrDefault(row);
    }

    /// <summary>A column's outline level: 0 outside any group, up to <see cref="MaxOutlineLevel"/>.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the sheet.</exception>
    public int ColumnOutlineLevel(int column)
    {
        CheckColumn(column);
        return _columnLevels.GetValueOrDefault(column);
    }

    /// <summary>Sets a row's outline level as it is, 0 to take it out of every group.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The row is outside the sheet, or the level outside 0 to 7.</exception>
    public void SetRowOutlineLevel(int row, int level)
    {
        CheckRow(row);
        ThrowIfCannotChange();
        SetLevel(_rowLevels, row, level);
    }

    /// <summary>Sets a column's outline level as it is, 0 to take it out of every group.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the sheet, or the level outside 0 to 7.</exception>
    public void SetColumnOutlineLevel(int column, int level)
    {
        CheckColumn(column);
        ThrowIfCannotChange();
        SetLevel(_columnLevels, column, level);
    }

    /// <summary>
    /// Groups rows, as Excel's Data &gt; Group does: each goes one level deeper. Their summary row is
    /// the one after them (or before, with <see cref="OutlineSummaryBelow"/> false).
    /// </summary>
    /// <param name="first">The first row.</param>
    /// <param name="last">The last row.</param>
    /// <exception cref="ArgumentOutOfRangeException">A row is outside the sheet, or first is after last.</exception>
    /// <exception cref="InvalidOperationException">A row is already at level 7.</exception>
    public void GroupRows(int first, int last) => Group(_rowLevels, first, last, rows: true, by: 1);

    /// <summary>Takes rows one level out, as Excel's Data &gt; Ungroup does; rows outside any group stay so.</summary>
    /// <exception cref="ArgumentOutOfRangeException">A row is outside the sheet, or first is after last.</exception>
    public void UngroupRows(int first, int last) => Group(_rowLevels, first, last, rows: true, by: -1);

    /// <summary>Groups columns, as Excel's Data &gt; Group does: each goes one level deeper.</summary>
    /// <exception cref="ArgumentOutOfRangeException">A column is outside the sheet, or first is after last.</exception>
    /// <exception cref="InvalidOperationException">A column is already at level 7.</exception>
    public void GroupColumns(int first, int last) => Group(_columnLevels, first, last, rows: false, by: 1);

    /// <summary>Takes columns one level out, as Excel's Data &gt; Ungroup does.</summary>
    /// <exception cref="ArgumentOutOfRangeException">A column is outside the sheet, or first is after last.</exception>
    public void UngroupColumns(int first, int last) => Group(_columnLevels, first, last, rows: false, by: -1);

    /// <summary>
    /// Collapses the group a row is in at its own level, as the group's minus button does: the
    /// group's rows are hidden and its summary row is marked collapsed. Groups inside it stay as they are.
    /// </summary>
    /// <param name="row">Any row of the group.</param>
    /// <exception cref="ArgumentOutOfRangeException">The row is outside the sheet.</exception>
    /// <exception cref="ArgumentException">The row is in no group.</exception>
    public void CollapseRowGroup(int row) => Collapse(row, rows: true, collapse: true);

    /// <summary>
    /// Expands the group a row is in at its own level, as the group's plus button does: its rows are
    /// shown, apart from those of groups inside it that stay collapsed.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The row is outside the sheet.</exception>
    /// <exception cref="ArgumentException">The row is in no group.</exception>
    public void ExpandRowGroup(int row) => Collapse(row, rows: true, collapse: false);

    /// <summary>Collapses the group a column is in at its own level, hiding its columns.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the sheet.</exception>
    /// <exception cref="ArgumentException">The column is in no group.</exception>
    public void CollapseColumnGroup(int column) => Collapse(column, rows: false, collapse: true);

    /// <summary>Expands the group a column is in at its own level, apart from collapsed groups inside it.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the sheet.</exception>
    /// <exception cref="ArgumentException">The column is in no group.</exception>
    public void ExpandColumnGroup(int column) => Collapse(column, rows: false, collapse: false);

    /// <summary>Whether the group a row is in at its own level is collapsed; false for a row in no group.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The row is outside the sheet.</exception>
    public bool IsRowGroupCollapsed(int row)
    {
        CheckRow(row);
        return Run(_rowLevels, row, rows: true) is { } run && _collapsedRows.Contains(Summary(run, rows: true));
    }

    /// <summary>Whether the group a column is in at its own level is collapsed; false for a column in no group.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the sheet.</exception>
    public bool IsColumnGroupCollapsed(int column)
    {
        CheckColumn(column);
        return Run(_columnLevels, column, rows: false) is { } run && _collapsedColumns.Contains(Summary(run, rows: false));
    }

    internal IReadOnlyDictionary<int, int> RowLevels => _rowLevels;

    internal IReadOnlyDictionary<int, int> ColumnLevels => _columnLevels;

    internal IReadOnlySet<int> CollapsedRows => _collapsedRows;

    internal IReadOnlySet<int> CollapsedColumns => _collapsedColumns;

    /// <summary>The deepest row level, as sheetFormatPr's outlineLevelRow gives it.</summary>
    internal int MaxRowLevel => _rowLevels.Count == 0 ? 0 : _rowLevels.Values.Max();

    internal int MaxColumnLevel => _columnLevels.Count == 0 ? 0 : _columnLevels.Values.Max();

    /// <summary>Sets what the file says, as read: levels kept in range, nothing hidden or shown.</summary>
    internal void LoadOutline(bool rows, int index, int level, bool collapsed)
    {
        if (level is > 0 and <= 7)
            (rows ? _rowLevels : _columnLevels)[index] = level;
        if (collapsed)
            (rows ? _collapsedRows : _collapsedColumns).Add(index);
    }

    private static void SetLevel(Dictionary<int, int> levels, int index, int level)
    {
        if (level is < 0 or > 7)
            throw new ArgumentOutOfRangeException(nameof(level), level, "Outline levels go from 0 to 7.");
        if (level == 0)
            levels.Remove(index);
        else
            levels[index] = level;
    }

    private void Group(Dictionary<int, int> levels, int first, int last, bool rows, int by)
    {
        if (rows)
        {
            CheckRow(first);
            CheckRow(last);
        }
        else
        {
            CheckColumn(first);
            CheckColumn(last);
        }

        if (first > last)
            throw new ArgumentOutOfRangeException(nameof(first), first, "The first must not come after the last.");
        ThrowIfCannotChange();
        if (by > 0 && Enumerable.Range(first, last - first + 1).Any(index => levels.GetValueOrDefault(index) >= MaxOutlineLevel))
            throw new InvalidOperationException("Excel groups to at most 7 levels.");
        for (var index = first; index <= last; index++)
            SetLevel(levels, index, Math.Max(0, levels.GetValueOrDefault(index) + by));
    }

    // The rows (or columns) of the group around an index at its level: the run with that level or deeper.
    private (int First, int Last, int Level)? Run(Dictionary<int, int> levels, int index, bool rows)
    {
        var level = levels.GetValueOrDefault(index);
        if (level == 0)
            return null;
        var max = rows ? CellAddress.MaxRow : CellAddress.MaxColumn;
        int first = index, last = index;
        while (first > 1 && levels.GetValueOrDefault(first - 1) >= level)
            first--;
        while (last < max && levels.GetValueOrDefault(last + 1) >= level)
            last++;
        return (first, last, level);
    }

    private int Summary((int First, int Last, int Level) run, bool rows)
    {
        var below = rows ? _summaryBelow : _summaryRight;
        var max = rows ? CellAddress.MaxRow : CellAddress.MaxColumn;
        return below ? Math.Min(run.Last + 1, max) : Math.Max(run.First - 1, 1);
    }

    private void Collapse(int index, bool rows, bool collapse)
    {
        if (rows)
            CheckRow(index);
        else
            CheckColumn(index);
        var levels = rows ? _rowLevels : _columnLevels;
        if (Run(levels, index, rows) is not { } run)
            throw new ArgumentException($"{(rows ? "Row" : "Column")} {index} is in no group.", nameof(index));
        ThrowIfCannotChange();
        var collapsed = rows ? _collapsedRows : _collapsedColumns;
        var summary = Summary(run, rows);
        if (collapse)
            collapsed.Add(summary);
        else
            collapsed.Remove(summary);

        // Expanding shows the group's rows, but not those of a collapsed group inside it.
        for (var i = run.First; i <= run.Last; i++)
        {
            var hidden = collapse || InCollapsedInner(i);
            if (rows)
                SetRowHidden(i, hidden);
            else
                SetColumnHidden(i, hidden);
        }

        bool InCollapsedInner(int i)
        {
            for (var level = run.Level + 1; level <= levels.GetValueOrDefault(i); level++)
            {
                int first = i, last = i;
                while (first > run.First && levels.GetValueOrDefault(first - 1) >= level)
                    first--;
                while (last < run.Last && levels.GetValueOrDefault(last + 1) >= level)
                    last++;
                if (collapsed.Contains(Summary((first, last, level), rows)))
                    return true;
            }

            return false;
        }
    }

    // Levels and collapsed marks follow inserted and deleted whole rows or columns; rows inserted
    // inside a group join it at the shallower of the levels around them, as in Excel.
    private void MoveOutline(SheetEdit edit)
    {
        if (!edit.Whole)
            return;
        var levels = edit.Rows ? _rowLevels : _columnLevels;
        var max = edit.Rows ? CellAddress.MaxRow : CellAddress.MaxColumn;
        MoveIndexes(levels, edit, max);
        MoveIndexes(edit.Rows ? _collapsedRows : _collapsedColumns, edit, max);
        if (edit.Delete || edit.At <= 1)
            return;
        var inside = Math.Min(levels.GetValueOrDefault(edit.At - 1), levels.GetValueOrDefault(edit.At + edit.Count));
        for (var index = edit.At; inside > 0 && index < edit.At + edit.Count && index <= max; index++)
            levels[index] = inside;
    }

    private void CopyOutlineFrom(Worksheet source)
    {
        foreach (var (row, level) in source._rowLevels)
            _rowLevels.Add(row, level);
        foreach (var (column, level) in source._columnLevels)
            _columnLevels.Add(column, level);
        _collapsedRows.UnionWith(source._collapsedRows);
        _collapsedColumns.UnionWith(source._collapsedColumns);
        _summaryBelow = source._summaryBelow;
        _summaryRight = source._summaryRight;
    }
}
