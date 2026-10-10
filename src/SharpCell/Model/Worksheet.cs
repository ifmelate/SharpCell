using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell;

/// <summary>
/// A sheet of a <see cref="SharpCell.Workbook"/>. Reading a cell that was never set costs nothing:
/// a <see cref="Cell"/> is a handle, and storage is allocated only when a value or formula is set.
/// </summary>
public sealed class Worksheet
{
    internal Worksheet(Workbook workbook, string name)
    {
        Workbook = workbook;
        Name = name;
    }

    /// <summary>The workbook the sheet belongs to.</summary>
    public Workbook Workbook { get; }

    /// <summary>The sheet name.</summary>
    public string Name { get; }

    internal SheetStore Store { get; } = new();

    /// <summary>A cell by A1 address such as <c>B3</c> (no <c>$</c>, no range).</summary>
    public Cell this[string address]
    {
        get
        {
            ArgumentNullException.ThrowIfNull(address);
            if (!CellAddress.TryParse(address, out var parsed))
                throw new ArgumentException($"'{address}' is not a cell address.", nameof(address));
            return new Cell(this, parsed.Row, parsed.Column);
        }
    }

    /// <summary>A cell by 1-based row and column.</summary>
    public Cell this[int row, int column]
    {
        get
        {
            if (row is < 1 or > CellAddress.MaxRow)
                throw new ArgumentOutOfRangeException(nameof(row), row, "Row is outside the sheet.");
            if (column is < 1 or > CellAddress.MaxColumn)
                throw new ArgumentOutOfRangeException(nameof(column), column, "Column is outside the sheet.");
            return new Cell(this, row, column);
        }
    }

    /// <summary>
    /// Makes a range a table that formulas can refer to by name, as in <c>Sales[Units]</c>. Column
    /// names come from the header row as it is now: a cell's value as text (a formula's too),
    /// <c>ColumnN</c> for an empty cell or an error (N counts from the table's first column), and a
    /// number added to a repeated name (<c>Units2</c>). Without a header row the columns are
    /// <c>Column1</c>, <c>Column2</c> and so on.
    /// Changing a header cell later does not rename its column.
    /// </summary>
    /// <param name="name">The table name; tables and defined names share one set of names, ignoring case.</param>
    /// <param name="range">An A1 range on this sheet, such as <c>A1:D10</c>, including the header and totals rows;
    /// a single cell such as <c>A1</c> is a one-cell table without a header row.</param>
    /// <param name="hasHeaderRow">Whether the first row holds the column names.</param>
    /// <param name="hasTotalsRow">Whether the last row holds totals.</param>
    /// <exception cref="ArgumentException">The name is not valid or is taken, the range is not an area,
    /// overlaps another table or has no row for data.</exception>
    public Table AddTable(string name, string range, bool hasHeaderRow = true, bool hasTotalsRow = false)
    {
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(range);
        // A single cell is a table of one data cell, without a header row.
        if (!ReferenceSyntax.TryParseA1Area(range, new CellAddress(1, 1), out var parsed) || parsed.Kind is not (AreaKind.Range or AreaKind.Cell))
            throw new ArgumentException($"'{range}' is not a range such as A1:D10.", nameof(range));
        var area = Area.Resolve(parsed, new CellAddress(1, 1));

        var columns = new string[area.Columns];
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < columns.Length; i++)
        {
            var header = hasHeaderRow ? HeaderText(area.FirstRow, area.FirstColumn + i) : "";
            var column = header.Length == 0 ? "Column" + (i + 1).ToString(CultureInfo.InvariantCulture) : header;
            var unique = column;
            for (var n = 2; !used.Add(unique); n++)
                unique = column + n.ToString(CultureInfo.InvariantCulture);
            columns[i] = unique;
        }

        return AddTable(name, area, hasHeaderRow, hasTotalsRow, columns);
    }

    /// <summary>Adds a table whose column names are known, as a file stores them.</summary>
    internal Table AddTable(string name, Area area, bool hasHeaderRow, bool hasTotalsRow, IReadOnlyList<string> columns)
    {
        if (!Workbook.IsValidName(name))
            throw new ArgumentException($"'{name}' is not a valid table name.", nameof(name));
        var rows = (hasHeaderRow ? 1 : 0) + 1 + (hasTotalsRow ? 1 : 0);
        if (area.Rows < rows)
            throw new ArgumentException($"The table needs at least {rows} rows: the header and totals rows it has and one row of data.", "range");
        if (columns.Count != area.Columns)
            throw new ArgumentException($"The table has {area.Columns} columns but {columns.Count} column names.", nameof(columns));
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in columns)
        {
            if (column.Length == 0 || !seen.Add(column))
                throw new ArgumentException($"Column names must be unique and not empty: '{column}'.", nameof(columns));
        }

        return Workbook.RegisterTable(new Table(this, name, area, hasHeaderRow, hasTotalsRow, columns));
    }

    // Excel turns a header formula into text when it makes the table; the formula's value counts,
    // calculated first when it is out of date.
    private string HeaderText(int row, int column)
    {
        var data = Store.Get(row, column);
        var value = data is { Formula: not null, IsDirty: true }
            ? Workbook.Calculation.EvaluateDetached(
                new ReferenceNode(null, AreaRef.Cell(new CellRef(AxisRef.Absolute(row), AxisRef.Absolute(column)))),
                this, new CellAddress(row, column), CancellationToken.None)
            : data?.Value ?? CellValue.Empty;
        var text = Coercion.ToText(value, Workbook.Culture);
        return text.Kind == CellValueKind.Text ? text.AsText() : "";
    }

    private readonly HashSet<int> _hiddenRows = [];

    /// <summary>The hidden rows, in no particular order.</summary>
    internal IReadOnlyCollection<int> HiddenRows => _hiddenRows;

    /// <summary>Whether a row is hidden.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The row is outside the sheet.</exception>
    public bool IsRowHidden(int row)
    {
        CheckRow(row);
        return _hiddenRows.Count != 0 && _hiddenRows.Contains(row);
    }

    /// <summary>
    /// Hides or shows a row. SUBTOTAL with codes 101–111 and AGGREGATE with options 1, 3, 5 and 7
    /// skip hidden rows; see <see cref="FilterMode"/> for SUBTOTAL 1–11. Like any change, it shows
    /// at the next <see cref="Workbook.Recalculate"/>.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The row is outside the sheet.</exception>
    public void SetRowHidden(int row, bool hidden)
    {
        CheckRow(row);
        if (hidden)
            _hiddenRows.Add(row);
        else
            _hiddenRows.Remove(row);
    }

    /// <summary>
    /// Whether the sheet has a filter with criteria, like <c>Worksheet.FilterMode</c> in Excel. Excel
    /// then treats every hidden row of the sheet as filtered out, so SUBTOTAL with codes 1–11 skips
    /// them as well; without a filter it counts them. AGGREGATE does not look at it.
    /// </summary>
    public bool FilterMode { get; set; }

    private static void CheckRow(int row)
    {
        if (row is < 1 or > CellAddress.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(row), row, "Row is outside the sheet.");
    }

    /// <summary>The sheet name.</summary>
    public override string ToString() => Name;
}
