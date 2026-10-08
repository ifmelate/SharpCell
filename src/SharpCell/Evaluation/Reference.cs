using System;
using System.Collections.Generic;

namespace SharpCell.Evaluation;

/// <summary>A normalized rectangle on a sheet: first corner is top-left. Rows and columns are 1-based.</summary>
internal readonly record struct Area(int FirstRow, int FirstColumn, int LastRow, int LastColumn)
{
    public int Rows => LastRow - FirstRow + 1;

    public int Columns => LastColumn - FirstColumn + 1;

    public bool IsSingleCell => FirstRow == LastRow && FirstColumn == LastColumn;

    public long CellCount => (long)Rows * Columns;

    public static Area Cell(int row, int column) => new(row, column, row, column);

    public static Area Bounding(Area a, Area b) => new(
        Math.Min(a.FirstRow, b.FirstRow), Math.Min(a.FirstColumn, b.FirstColumn),
        Math.Max(a.LastRow, b.LastRow), Math.Max(a.LastColumn, b.LastColumn));

    public bool Contains(int row, int column) =>
        row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;

    public bool TryIntersect(Area other, out Area result)
    {
        result = new Area(
            Math.Max(FirstRow, other.FirstRow), Math.Max(FirstColumn, other.FirstColumn),
            Math.Min(LastRow, other.LastRow), Math.Min(LastColumn, other.LastColumn));
        return result.FirstRow <= result.LastRow && result.FirstColumn <= result.LastColumn;
    }

    /// <summary>Resolves a parsed reference at the formula's cell; corners are ordered.</summary>
    public static Area Resolve(AreaRef area, CellAddress origin)
    {
        switch (area.Kind)
        {
            case AreaKind.Cell:
                var cell = area.First.Resolve(origin);
                return Cell(cell.Row, cell.Column);
            case AreaKind.Range:
                var first = area.First.Resolve(origin);
                var last = area.Last.Resolve(origin);
                return Bounding(Cell(first.Row, first.Column), Cell(last.Row, last.Column));
            case AreaKind.Rows:
                var r1 = area.First.Row.Resolve(origin.Row, CellAddress.MaxRow);
                var r2 = area.Last.Row.Resolve(origin.Row, CellAddress.MaxRow);
                return new Area(Math.Min(r1, r2), 1, Math.Max(r1, r2), CellAddress.MaxColumn);
            default:
                var c1 = area.First.Column.Resolve(origin.Column, CellAddress.MaxColumn);
                var c2 = area.Last.Column.Resolve(origin.Column, CellAddress.MaxColumn);
                return new Area(1, Math.Min(c1, c2), CellAddress.MaxRow, Math.Max(c1, c2));
        }
    }
}

internal readonly record struct SheetArea(Worksheet Sheet, Area Area);

/// <summary>
/// A reference produced during evaluation: one or more areas, possibly on several sheets (3D).
/// References are never stored in cells; they are turned into values only when an operator or
/// function needs values.
/// </summary>
internal sealed class Reference(IReadOnlyList<SheetArea> areas)
{
    public Reference(Worksheet sheet, Area area)
        : this([new SheetArea(sheet, area)])
    {
    }

    public IReadOnlyList<SheetArea> Areas { get; } = areas;

    public bool IsSingleArea => Areas.Count == 1;
}
