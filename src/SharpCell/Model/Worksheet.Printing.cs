using System;
using System.Collections.Generic;
using System.Linq;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell;

// How the sheet prints: orientation, margins, the print area, rows and columns repeated on every
// page, and manual page breaks, as Excel's Page Layout tab sets them. Calculation never sees them.
public sealed partial class Worksheet
{
    private PageOrientation _orientation;
    private PageMargins _margins = PageMargins.Normal;
    private readonly List<Area> _printArea = [];
    private (int First, int Last)? _titleRows;
    private (int First, int Last)? _titleColumns;
    private readonly SortedSet<int> _rowBreaks = [];
    private readonly SortedSet<int> _columnBreaks = [];

    /// <summary>How the sheet is turned on the printed page; <see cref="SharpCell.PageOrientation.Portrait"/> unless set otherwise.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not a defined orientation.</exception>
    public PageOrientation PageOrientation
    {
        get => _orientation;
        set
        {
            if (value is not (PageOrientation.Portrait or PageOrientation.Landscape))
                throw new ArgumentOutOfRangeException(nameof(value), value, "Not a page orientation.");
            Workbook.ThrowIfInCustomFunction();
            _orientation = value;
        }
    }

    /// <summary>The margins of the printed pages; <see cref="SharpCell.PageMargins.Normal"/> unless set otherwise.</summary>
    /// <exception cref="ArgumentNullException">The value is null.</exception>
    public PageMargins PageMargins
    {
        get => _margins;
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            Workbook.ThrowIfInCustomFunction();
            _margins = value;
        }
    }

    /// <summary>
    /// The cells that print, as Excel's Print Area sets them: one range such as <c>A1:F40</c>, or
    /// several separated by commas, each printed on pages of its own. Null prints every cell that
    /// holds something, which is also what setting null gives. Read without <c>$</c> signs; written
    /// with or without them. The ranges move with their cells when rows, columns or cells are
    /// inserted or deleted, as the defined name Excel keeps them in does.
    /// </summary>
    /// <exception cref="ArgumentException">The text is not A1 ranges of this sheet separated by commas.</exception>
    public string? PrintArea
    {
        get => _printArea.Count == 0 ? null : string.Join(",", _printArea.Select(area => new CellRange(this, area).Address));
        set
        {
            var areas = value is null ? [] : ParseAreas(value, nameof(value));
            Workbook.ThrowIfInCustomFunction();
            _printArea.Clear();
            _printArea.AddRange(areas);
        }
    }

    /// <summary>
    /// The rows printed at the top of every page, as Excel's Print Titles sets them: <c>1:1</c> or
    /// <c>1:3</c>. Null for none. They move with their rows when rows are inserted or deleted.
    /// </summary>
    /// <exception cref="ArgumentException">The text is not a span of whole rows such as <c>1:2</c>.</exception>
    public string? PrintTitleRows
    {
        get => _titleRows is { } rows ? $"{rows.First}:{rows.Last}" : null;
        set
        {
            (int, int)? rows = value is null ? null : ParseLines(value, AreaKind.Rows, nameof(value));
            Workbook.ThrowIfInCustomFunction();
            _titleRows = rows;
        }
    }

    /// <summary>
    /// The columns printed at the left of every page: <c>A:A</c> or <c>A:B</c>. Null for none. They
    /// move with their columns when columns are inserted or deleted.
    /// </summary>
    /// <exception cref="ArgumentException">The text is not a span of whole columns such as <c>A:B</c>.</exception>
    public string? PrintTitleColumns
    {
        get => _titleColumns is { } columns ? $"{ColumnNames.ToName(columns.First)}:{ColumnNames.ToName(columns.Last)}" : null;
        set
        {
            (int, int)? columns = value is null ? null : ParseLines(value, AreaKind.Columns, nameof(value));
            Workbook.ThrowIfInCustomFunction();
            _titleColumns = columns;
        }
    }

    /// <summary>The rows that start a new printed page, as Excel's Insert Page Break puts them, in order.</summary>
    public IReadOnlyList<int> RowPageBreaks => [.. _rowBreaks];

    /// <summary>The columns that start a new printed page, in order.</summary>
    public IReadOnlyList<int> ColumnPageBreaks => [.. _columnBreaks];

    /// <summary>Starts a new printed page at a row: a page break above it. Breaks move with their rows.</summary>
    /// <param name="row">The first row of the new page, from 2.</param>
    /// <exception cref="ArgumentOutOfRangeException">The row is not on the sheet or is row 1.</exception>
    public void AddRowPageBreak(int row)
    {
        if (row is < 2 or > CellAddress.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(row), row, "A page break goes above a row from 2 to the last.");
        Workbook.ThrowIfInCustomFunction();
        _rowBreaks.Add(row);
    }

    /// <summary>Removes the page break above a row, if there is one.</summary>
    /// <param name="row">The row.</param>
    public void RemoveRowPageBreak(int row)
    {
        Workbook.ThrowIfInCustomFunction();
        _rowBreaks.Remove(row);
    }

    /// <summary>Starts a new printed page at a column: a page break left of it.</summary>
    /// <param name="column">The first column of the new page, from 2 (B).</param>
    /// <exception cref="ArgumentOutOfRangeException">The column is not on the sheet or is column A.</exception>
    public void AddColumnPageBreak(int column)
    {
        if (column is < 2 or > CellAddress.MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(column), column, "A page break goes left of a column from B to the last.");
        Workbook.ThrowIfInCustomFunction();
        _columnBreaks.Add(column);
    }

    /// <summary>Removes the page break left of a column, if there is one.</summary>
    /// <param name="column">The column.</param>
    public void RemoveColumnPageBreak(int column)
    {
        Workbook.ThrowIfInCustomFunction();
        _columnBreaks.Remove(column);
    }

    /// <summary>The print area's ranges; empty for none.</summary>
    internal IReadOnlyList<Area> PrintAreas => _printArea;

    internal (int First, int Last)? TitleRows => _titleRows;

    internal (int First, int Last)? TitleColumns => _titleColumns;

    /// <summary>Sets what the file says about printing, as read: nothing checks or moves here.</summary>
    internal void LoadPrinting(IEnumerable<Area>? printArea, (int, int)? titleRows, (int, int)? titleColumns)
    {
        if (printArea is not null)
        {
            _printArea.Clear();
            _printArea.AddRange(printArea);
        }

        _titleRows = titleRows ?? _titleRows;
        _titleColumns = titleColumns ?? _titleColumns;
    }

    internal void LoadPageBreaks(IEnumerable<int> rows, IEnumerable<int> columns)
    {
        _rowBreaks.UnionWith(rows.Where(row => row is >= 2 and <= CellAddress.MaxRow));
        _columnBreaks.UnionWith(columns.Where(column => column is >= 2 and <= CellAddress.MaxColumn));
    }

    // Print settings follow inserted and deleted rows and columns: the print area as a reference to
    // its cells does, titles and page breaks only with whole rows or columns.
    private void MovePrinting(SheetEdit edit)
    {
        for (var i = _printArea.Count - 1; i >= 0; i--)
        {
            if (edit.Map(_printArea[i]) is { } moved)
                _printArea[i] = moved;
            else
                _printArea.RemoveAt(i);
        }

        if (!edit.Whole)
            return;
        var max = edit.Rows ? CellAddress.MaxRow : CellAddress.MaxColumn;
        if (edit.Rows)
            _titleRows = _titleRows is { } rows ? edit.MapSpan(rows.First, rows.Last, max) : null;
        else
            _titleColumns = _titleColumns is { } columns ? edit.MapSpan(columns.First, columns.Last, max) : null;
        var breaks = edit.Rows ? _rowBreaks : _columnBreaks;
        var kept = breaks.Select(index => edit.Map(index, max)).OfType<int>().Where(index => index >= 2).ToList();
        breaks.Clear();
        breaks.UnionWith(kept);
    }

    private void CopyPrintingFrom(Worksheet source)
    {
        _orientation = source._orientation;
        _margins = source._margins;
        _printArea.AddRange(source._printArea);
        _titleRows = source._titleRows;
        _titleColumns = source._titleColumns;
        _rowBreaks.UnionWith(source._rowBreaks);
        _columnBreaks.UnionWith(source._columnBreaks);
    }

    private static List<Area> ParseAreas(string text, string parameter)
    {
        var areas = new List<Area>();
        foreach (var part in text.Split(','))
        {
            if (!ReferenceSyntax.TryParseA1Area(part.Trim().AsSpan(), new CellAddress(1, 1), out var area))
                throw new ArgumentException($"'{text}' is not A1 ranges separated by commas, such as A1:F40.", parameter);
            areas.Add(Area.Resolve(area, new CellAddress(1, 1)));
        }

        return areas;
    }

    private static (int, int) ParseLines(string text, AreaKind kind, string parameter)
    {
        if (!ReferenceSyntax.TryParseA1Area(text.Trim().AsSpan(), new CellAddress(1, 1), out var parsed) || parsed.Kind != kind)
            throw new ArgumentException(kind == AreaKind.Rows ? $"'{text}' is not whole rows such as 1:2." : $"'{text}' is not whole columns such as A:B.", parameter);
        var area = Area.Resolve(parsed, new CellAddress(1, 1));
        return kind == AreaKind.Rows ? (area.FirstRow, area.LastRow) : (area.FirstColumn, area.LastColumn);
    }
}
