namespace SharpCell;

/// <summary>
/// One coordinate of a reference: an absolute 1-based index (<c>$A</c>, <c>R1</c>) or an offset
/// from the cell that holds the formula (<c>A</c>, <c>R[-1]</c>). Storing offsets lets one formula
/// tree serve every cell of a shared or copied formula.
/// </summary>
internal readonly record struct AxisRef(int Value, bool IsAbsolute)
{
    public static AxisRef Absolute(int index) => new(index, true);

    public static AxisRef Relative(int offset) => new(offset, false);

    /// <summary>Resolves to a 1-based index. Relative positions wrap around the sheet edge, as in Excel.</summary>
    public int Resolve(int origin, int max)
    {
        if (IsAbsolute)
            return Value;

        var zeroBased = (origin - 1 + Value) % max;
        return (zeroBased < 0 ? zeroBased + max : zeroBased) + 1;
    }
}

internal readonly record struct CellRef(AxisRef Row, AxisRef Column)
{
    public CellAddress Resolve(CellAddress origin) =>
        new(Row.Resolve(origin.Row, CellAddress.MaxRow), Column.Resolve(origin.Column, CellAddress.MaxColumn));
}

internal enum AreaKind
{
    Cell,
    Range,
    Rows,
    Columns,
}

/// <summary>
/// A rectangular reference without a sheet: a cell, a cell range, whole rows or whole columns.
/// Corners are kept as written (<c>B2:A1</c> stays <c>B2:A1</c>); evaluation normalizes them.
/// </summary>
internal readonly record struct AreaRef(AreaKind Kind, CellRef First, CellRef Last)
{
    public static AreaRef Cell(CellRef cell) => new(AreaKind.Cell, cell, cell);

    public static AreaRef Range(CellRef first, CellRef last) => new(AreaKind.Range, first, last);

    public static AreaRef Rows(AxisRef first, AxisRef last) => new(AreaKind.Rows, new CellRef(first, default), new CellRef(last, default));

    public static AreaRef Columns(AxisRef first, AxisRef last) => new(AreaKind.Columns, new CellRef(default, first), new CellRef(default, last));
}
