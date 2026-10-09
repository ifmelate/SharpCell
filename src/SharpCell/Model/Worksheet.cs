using System;
using System.Collections.Generic;
using System.Globalization;
using SharpCell.Evaluation;

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
    /// names come from the header row as it is now: a cell's text, <c>ColumnN</c> for an empty cell
    /// (N counts from the table's first column) and a number added to a repeated name
    /// (<c>Units2</c>). Without a header row the columns are <c>Column1</c>, <c>Column2</c> and so on.
    /// Changing a header cell later does not rename its column.
    /// </summary>
    /// <param name="name">The table name; tables and defined names share one set of names, ignoring case.</param>
    /// <param name="range">An A1 range on this sheet, such as <c>A1:D10</c>, including the header and totals rows.</param>
    /// <param name="hasHeaderRow">Whether the first row holds the column names.</param>
    /// <param name="hasTotalsRow">Whether the last row holds totals.</param>
    /// <exception cref="ArgumentException">The name is not valid or is taken, the range is not an area,
    /// overlaps another table or has no row for data.</exception>
    public Table AddTable(string name, string range, bool hasHeaderRow = true, bool hasTotalsRow = false)
    {
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(range);
        if (!ReferenceSyntax.TryParseA1Area(range, new CellAddress(1, 1), out var parsed) || parsed.Kind != AreaKind.Range)
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

    private string HeaderText(int row, int column)
    {
        var value = Store.Get(row, column)?.Value ?? CellValue.Empty;
        var text = Coercion.ToText(value, Workbook.Culture);
        return text.Kind == CellValueKind.Text ? text.AsText() : "";
    }

    /// <summary>The sheet name.</summary>
    public override string ToString() => Name;
}
