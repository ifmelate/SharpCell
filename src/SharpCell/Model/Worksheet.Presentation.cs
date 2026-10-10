using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell;

// How the sheet looks rather than what it calculates: cell styles, column widths, row heights,
// hidden columns, merged cells, frozen panes and gridlines. None of it is cell content, and calculation never sees it.
public sealed partial class Worksheet
{
    private const double MaxColumnWidth = 255;
    private const double MaxRowHeight = 409;

    private readonly Dictionary<CellAddress, CellStyle> _styles = [];
    private readonly Dictionary<int, double> _columnWidths = [];
    private readonly Dictionary<int, double> _rowHeights = [];
    private readonly HashSet<int> _hiddenColumns = [];
    private readonly List<Area> _merged = [];
    private double? _defaultColumnWidth;
    private double? _defaultRowHeight;
    private int _frozenRows;
    private int _frozenColumns;
    private bool _showGridlines = true;

    /// <summary>
    /// How many rows at the top stay in place while the rest scrolls, as Excel's Freeze Panes
    /// keeps them; 0 for none.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The count is negative or leaves no row to scroll.</exception>
    public int FrozenRows
    {
        get => _frozenRows;
        set
        {
            if (value is < 0 or >= CellAddress.MaxRow)
                throw new ArgumentOutOfRangeException(nameof(value), value, "Frozen rows must leave at least one row to scroll.");
            Workbook.ThrowIfInCustomFunction();
            _frozenRows = value;
        }
    }

    /// <summary>How many columns on the left stay in place while the rest scrolls; 0 for none.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The count is negative or leaves no column to scroll.</exception>
    public int FrozenColumns
    {
        get => _frozenColumns;
        set
        {
            if (value is < 0 or >= CellAddress.MaxColumn)
                throw new ArgumentOutOfRangeException(nameof(value), value, "Frozen columns must leave at least one column to scroll.");
            Workbook.ThrowIfInCustomFunction();
            _frozenColumns = value;
        }
    }

    /// <summary>Whether Excel draws gridlines between the sheet's cells; true unless turned off.</summary>
    public bool ShowGridlines
    {
        get => _showGridlines;
        set
        {
            Workbook.ThrowIfInCustomFunction();
            _showGridlines = value;
        }
    }

    internal CellStyle? StyleAt(int row, int column) =>
        _styles.Count != 0 && _styles.TryGetValue(new CellAddress(row, column), out var style) ? style : null;

    /// <summary>Sets or (with null) removes the style of a cell.</summary>
    internal void SetStyle(int row, int column, CellStyle? style)
    {
        if (style is null)
            _styles.Remove(new CellAddress(row, column));
        else
            _styles[new CellAddress(row, column)] = style;
    }

    /// <summary>The cells with a style of their own, in no particular order.</summary>
    internal IReadOnlyDictionary<CellAddress, CellStyle> Styles => _styles;

    internal IReadOnlyDictionary<int, double> ColumnWidths => _columnWidths;

    internal IReadOnlyDictionary<int, double> RowHeights => _rowHeights;

    internal IReadOnlyCollection<int> HiddenColumns => _hiddenColumns;

    /// <summary>
    /// The width of columns without one of their own, as .xlsx files store it: in characters of the
    /// widest digit of the workbook's default font, padding included. Excel's standard 8.43 characters
    /// of Calibri 11 are 9.140625 here. Null when not given; Excel then uses 8 characters.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The width is outside 0 to 255.</exception>
    public double? DefaultColumnWidth
    {
        get => _defaultColumnWidth;
        set
        {
            Workbook.ThrowIfInCustomFunction();
            _defaultColumnWidth = value is { } width ? CheckSize(width, MaxColumnWidth, nameof(value)) : null;
        }
    }

    /// <summary>
    /// The height of rows without one of their own, in points. Null when not given; Excel then
    /// fits it to the default font, 15 points for Calibri 11.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The height is outside 0 to 409.</exception>
    public double? DefaultRowHeight
    {
        get => _defaultRowHeight;
        set
        {
            Workbook.ThrowIfInCustomFunction();
            _defaultRowHeight = value is { } height ? CheckSize(height, MaxRowHeight, nameof(value)) : null;
        }
    }

    /// <summary>A column's own width, in the units of <see cref="DefaultColumnWidth"/>; null when it has none.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the sheet.</exception>
    public double? ColumnWidth(int column)
    {
        CheckColumn(column);
        return _columnWidths.Count != 0 && _columnWidths.TryGetValue(column, out var width) ? width : null;
    }

    /// <summary>Gives a column a width of its own, in the units of <see cref="DefaultColumnWidth"/>; null takes it away.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the sheet, or the width outside 0 to 255.</exception>
    public void SetColumnWidth(int column, double? width)
    {
        CheckColumn(column);
        Workbook.ThrowIfInCustomFunction();
        if (width is { } value)
            _columnWidths[column] = CheckSize(value, MaxColumnWidth, nameof(width));
        else
            _columnWidths.Remove(column);
    }

    /// <summary>A row's own height in points; null when it has none.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The row is outside the sheet.</exception>
    public double? RowHeight(int row)
    {
        CheckRow(row);
        return _rowHeights.Count != 0 && _rowHeights.TryGetValue(row, out var height) ? height : null;
    }

    /// <summary>Gives a row a height of its own in points; null takes it away.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The row is outside the sheet, or the height outside 0 to 409.</exception>
    public void SetRowHeight(int row, double? height)
    {
        CheckRow(row);
        Workbook.ThrowIfInCustomFunction();
        if (height is { } value)
            _rowHeights[row] = CheckSize(value, MaxRowHeight, nameof(height));
        else
            _rowHeights.Remove(row);
    }

    /// <summary>Whether a column is hidden.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the sheet.</exception>
    public bool IsColumnHidden(int column)
    {
        CheckColumn(column);
        return _hiddenColumns.Count != 0 && _hiddenColumns.Contains(column);
    }

    /// <summary>Hides or shows a column. Unlike hidden rows, hidden columns change no result, as in Excel.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the sheet.</exception>
    public void SetColumnHidden(int column, bool hidden)
    {
        CheckColumn(column);
        Workbook.ThrowIfInCustomFunction();
        if (hidden)
            _hiddenColumns.Add(column);
        else
            _hiddenColumns.Remove(column);
    }

    /// <summary>The merged areas, in the order they were merged.</summary>
    public IReadOnlyList<CellRange> MergedAreas => _merged.ConvertAll(area => new CellRange(this, area));

    internal IReadOnlyList<Area> MergedAreaList => _merged;

    /// <summary>
    /// Merges a range of at least two cells into one, as Excel shows it: the top-left cell covers
    /// the range. The values of the other cells are kept; Excel's Merge command would clear them.
    /// </summary>
    /// <exception cref="ArgumentException">The address is not a range of two cells or more, or it overlaps a merged area.</exception>
    public void Merge(string address)
    {
        var area = Range(address).Area;
        Workbook.ThrowIfInCustomFunction();
        if (area.IsSingleCell)
            throw new ArgumentException($"'{address}' is one cell; a merge needs two or more.", nameof(address));
        if (!TryMerge(area))
            throw new ArgumentException($"'{address}' overlaps a merged area.", nameof(address));
    }

    /// <summary>Undoes a merge; the address must be exactly a merged area.</summary>
    /// <exception cref="ArgumentException">The address is not a merged area.</exception>
    public void Unmerge(string address)
    {
        var area = Range(address).Area;
        Workbook.ThrowIfInCustomFunction();
        if (!_merged.Remove(area))
            throw new ArgumentException($"'{address}' is not a merged area.", nameof(address));
    }

    /// <summary>Adds a merged area unless it overlaps one; the reader skips such overlaps in damaged files.</summary>
    internal bool TryMerge(Area area)
    {
        foreach (var merged in _merged)
        {
            if (merged.TryIntersect(area, out _))
                return false;
        }

        _merged.Add(area);
        return true;
    }

    private void CopyPresentationFrom(Worksheet source)
    {
        foreach (var (address, style) in source._styles)
            _styles.Add(address, style);
        foreach (var (column, width) in source._columnWidths)
            _columnWidths.Add(column, width);
        foreach (var (row, height) in source._rowHeights)
            _rowHeights.Add(row, height);
        _hiddenColumns.UnionWith(source._hiddenColumns);
        _merged.AddRange(source._merged);
        _defaultColumnWidth = source._defaultColumnWidth;
        _defaultRowHeight = source._defaultRowHeight;
        _frozenRows = source._frozenRows;
        _frozenColumns = source._frozenColumns;
        _showGridlines = source._showGridlines;
    }

    private static double CheckSize(double size, double max, string parameter)
    {
        if (!(size >= 0 && size <= max))
            throw new ArgumentOutOfRangeException(parameter, size, $"The size must be from 0 to {max}.");
        return size;
    }

    private static void CheckColumn(int column)
    {
        if (column is < 1 or > CellAddress.MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(column), column, "Column is outside the sheet.");
    }
}
