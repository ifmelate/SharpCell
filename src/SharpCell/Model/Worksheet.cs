using System;

namespace SharpCell;

public sealed class Worksheet
{
    internal Worksheet(Workbook workbook, string name)
    {
        Workbook = workbook;
        Name = name;
    }

    public Workbook Workbook { get; }

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

    public override string ToString() => Name;
}
