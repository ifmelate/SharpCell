using System;
using System.Collections.Generic;
using System.Linq;
using SharpCell.Evaluation;

namespace SharpCell;

/// <summary>
/// The filter of a sheet, Excel's Data &gt; Filter: buttons in the first row of its range, and
/// criteria on some of its columns that hide the rows below not meeting them. Turn it on with
/// <see cref="Worksheet.SetAutoFilter"/>.
/// </summary>
public sealed class AutoFilter
{
    private readonly SortedDictionary<int, ColumnFilter> _columns = [];

    internal AutoFilter(Worksheet sheet, Area area)
    {
        Worksheet = sheet;
        Area = area;
    }

    /// <summary>The sheet the filter is on.</summary>
    public Worksheet Worksheet { get; }

    internal Area Area { get; set; }

    /// <summary>The filtered range, its first row holding the buttons.</summary>
    public CellRange Range => new(Worksheet, Area);

    /// <summary>The criteria by sheet column (A is 1), in column order.</summary>
    public IReadOnlyDictionary<int, ColumnFilter> Columns => _columns;

    /// <summary>
    /// What applying worked out for each column, for saving it as Excel does: the average, the
    /// lowest number shown by a top filter, the first and the day after the last date shown.
    /// </summary>
    internal Dictionary<int, (double Value, double? Max)> Limits { get; } = [];

    /// <summary>
    /// Filters a column and applies the filter at once, as choosing criteria in Excel's menu does:
    /// rows of the range below its first row show when they meet the criteria of every column, and
    /// hide otherwise. Formulas out of date in the range are calculated first.
    /// </summary>
    /// <param name="column">The sheet column (A is 1), inside the range.</param>
    /// <param name="filter">The criteria, replacing the column's.</param>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the range.</exception>
    /// <exception cref="InvalidOperationException">The filter was removed, or its sheet was.</exception>
    /// <exception cref="NotSupportedException">Another column filters by colour, which SharpCell cannot apply.</exception>
    public void Filter(int column, ColumnFilter filter)
    {
        ArgumentNullException.ThrowIfNull(filter);
        CheckColumn(column);
        CheckLive();
        _columns[column] = filter;
        Apply();
    }

    /// <summary>Takes the criteria off a column and applies the filter again.</summary>
    /// <inheritdoc cref="Filter" path="/exception"/>
    public void Clear(int column)
    {
        CheckColumn(column);
        CheckLive();
        _columns.Remove(column);
        Apply();
    }

    /// <summary>Takes the criteria off every column, showing every row of the range, as Excel's Clear does. The buttons stay.</summary>
    /// <exception cref="InvalidOperationException">The filter was removed, or its sheet was.</exception>
    public void Clear()
    {
        CheckLive();
        _columns.Clear();
        Apply();
    }

    /// <summary>
    /// Applies the criteria again, as Excel's Reapply does after the data changed: each row of the
    /// range below its first row shows when it meets the criteria of every column, and hides
    /// otherwise, whether or not it was hidden before. Formulas out of date in the range are
    /// calculated first.
    /// </summary>
    /// <exception cref="InvalidOperationException">The filter was removed, or its sheet was.</exception>
    /// <exception cref="NotSupportedException">A column filters by colour, which SharpCell cannot apply.</exception>
    public void Apply()
    {
        CheckLive();
        if (_columns.FirstOrDefault(c => c.Value.Type == ColumnFilterType.ByColor) is { Value: not null } colored)
            throw new NotSupportedException($"Column {new CellAddress(1, colored.Key).ToString()[..^1]} is filtered by colour, which SharpCell cannot apply.");
        Worksheet.ApplyFilter(this);
    }

    /// <summary>Sets a column's criteria as a file has them, without applying.</summary>
    internal void Load(int column, ColumnFilter filter) => _columns[column] = filter;

    /// <summary>A copy for another sheet: a clone of the workbook.</summary>
    internal AutoFilter CopyFor(Worksheet sheet)
    {
        var copy = new AutoFilter(sheet, Area);
        foreach (var (column, filter) in _columns)
            copy._columns[column] = filter;
        foreach (var (column, limit) in Limits)
            copy.Limits[column] = limit;
        return copy;
    }

    /// <summary>Moves the criteria with their columns; those of deleted columns go.</summary>
    internal void MoveColumns(Func<int, int?> map)
    {
        var moved = _columns.Select(pair => (Column: map(pair.Key), pair.Value)).Where(pair => pair.Column is not null).ToList();
        _columns.Clear();
        foreach (var (column, filter) in moved)
            _columns[column!.Value] = filter;
        Limits.Clear();
    }

    private void CheckColumn(int column)
    {
        if (column < Area.FirstColumn || column > Area.LastColumn)
            throw new ArgumentOutOfRangeException(nameof(column), column, "The column is outside the filtered range.");
    }

    private void CheckLive()
    {
        if (Worksheet.AutoFilter != this)
            throw new InvalidOperationException("The filter was removed from its sheet.");
        Worksheet.ThrowIfCannotChange();
    }
}
