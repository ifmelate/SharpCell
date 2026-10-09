using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell;

/// <summary>
/// An Excel table: a named range with column names, an optional header row and an optional totals
/// row. Formulas refer to it by name, as in <c>Sales[Units]</c>. A table does not change after it is
/// created.
/// </summary>
public sealed class Table
{
    private readonly Dictionary<string, int> _columnIndex = new(StringComparer.OrdinalIgnoreCase);

    internal Table(Worksheet worksheet, string name, Area area, bool hasHeaderRow, bool hasTotalsRow, IReadOnlyList<string> columns)
    {
        Worksheet = worksheet;
        Name = name;
        Area = area;
        HasHeaderRow = hasHeaderRow;
        HasTotalsRow = hasTotalsRow;
        Columns = columns;
        for (var i = 0; i < columns.Count; i++)
            _columnIndex[columns[i]] = i;
    }

    /// <summary>The name formulas use, such as <c>Sales</c> in <c>Sales[Units]</c>.</summary>
    public string Name { get; }

    /// <summary>The sheet the table is on.</summary>
    public Worksheet Worksheet { get; }

    /// <summary>The whole table as an A1 range, header and totals rows included, such as <c>A1:D10</c>.</summary>
    public string Range => $"{new CellAddress(Area.FirstRow, Area.FirstColumn)}:{new CellAddress(Area.LastRow, Area.LastColumn)}";

    /// <summary>Whether the first row holds the column names.</summary>
    public bool HasHeaderRow { get; }

    /// <summary>Whether the last row holds totals.</summary>
    public bool HasTotalsRow { get; }

    /// <summary>The column names from left to right.</summary>
    public IReadOnlyList<string> Columns { get; }

    internal Area Area { get; }

    internal int FirstDataRow => Area.FirstRow + (HasHeaderRow ? 1 : 0);

    internal int LastDataRow => Area.LastRow - (HasTotalsRow ? 1 : 0);

    /// <summary>The 0-based position of a column, ignoring case; -1 when the table has none of that name.</summary>
    internal int ColumnIndex(string name) => _columnIndex.TryGetValue(name, out var index) ? index : -1;

    /// <summary>The table name.</summary>
    public override string ToString() => Name;
}
