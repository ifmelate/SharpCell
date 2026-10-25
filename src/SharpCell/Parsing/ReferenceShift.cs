using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Parsing;

/// <summary>
/// Rows or columns inserted into or deleted from a sheet: <paramref name="Count"/> of them at
/// <paramref name="At"/> (1-based). With a <paramref name="Band"/>, only the cells in those columns
/// (for rows) or rows (for columns) move, as Excel's "Shift cells down" or "Shift cells left" does.
/// </summary>
/// <param name="Rows">Rows; otherwise columns.</param>
/// <param name="At">The first row or column inserted or deleted.</param>
/// <param name="Count">How many.</param>
/// <param name="Delete">Deleted; otherwise inserted.</param>
/// <param name="Band">The columns (for rows) or rows (for columns) whose cells move; null for whole rows or columns.</param>
internal readonly record struct SheetEdit(bool Rows, int At, int Count, bool Delete, (int First, int Last)? Band = null)
{
    /// <summary>Whether whole rows or columns move, not only the cells of a band.</summary>
    public bool Whole => Band is null;

    /// <summary>Whether a column (for rows) or row (for columns) is in the band.</summary>
    public bool InBand(int across) => Band is not { } band || (across >= band.First && across <= band.Last);

    /// <summary>Whether a span across the edit lies wholly in the band.</summary>
    public bool Covers(int first, int last) => Band is not { } band || (first >= band.First && last <= band.Last);

    /// <summary>Whether a span across the edit lies wholly outside the band.</summary>
    public bool Misses(int first, int last) => Band is { } band && (last < band.First || first > band.Last);

    /// <summary>Where a row or column goes; null when it is deleted or pushed past the sheet's end.</summary>
    public int? Map(int index, int max)
    {
        if (!Delete)
            return index < At ? index : index + Count <= max ? index + Count : null;
        return index < At ? index : index >= At + Count ? index - Count : null;
    }

    /// <summary>Where a cell goes; null when it is deleted or pushed past the sheet's end.</summary>
    public CellAddress? Map(CellAddress cell)
    {
        if (Rows)
            return !InBand(cell.Column) ? cell : Map(cell.Row, CellAddress.MaxRow) is { } r ? new CellAddress(r, cell.Column) : null;
        return !InBand(cell.Row) ? cell : Map(cell.Column, CellAddress.MaxColumn) is { } c ? new CellAddress(cell.Row, c) : null;
    }

    /// <summary>
    /// Where a span of rows or columns goes, as a range does: one inserted strictly inside it grows
    /// it, one at its start moves it; deleting part of it shrinks it. Null when all of it is deleted
    /// or pushed past the end.
    /// </summary>
    public (int First, int Last)? MapSpan(int first, int last, int max)
    {
        if (!Delete)
        {
            if (first >= At)
                return first + Count > max ? null : (first + Count, Math.Min(last + Count, max));
            return last >= At ? (first, Math.Min(last + Count, max)) : (first, last);
        }

        var end = At + Count - 1;
        if (first >= At && last <= end)
            return null;
        var newFirst = first < At ? first : first > end ? first - Count : At;
        var newLast = last < At ? last : last > end ? last - Count : At - 1;
        return (newFirst, newLast);
    }

    /// <summary>
    /// Where an area goes along the edited direction, as a reference to it does; null when it is
    /// gone. An area that is not wholly in the band stays where it is, as Excel leaves a reference
    /// that only some of the moved cells are in.
    /// </summary>
    public Area? Map(Area area)
    {
        if (Rows)
        {
            if (!Covers(area.FirstColumn, area.LastColumn))
                return area;
            return MapSpan(area.FirstRow, area.LastRow, CellAddress.MaxRow) is { } rows
                ? area with { FirstRow = rows.First, LastRow = rows.Last }
                : null;
        }

        if (!Covers(area.FirstRow, area.LastRow))
            return area;
        return MapSpan(area.FirstColumn, area.LastColumn, CellAddress.MaxColumn) is { } columns
            ? area with { FirstColumn = columns.First, LastColumn = columns.Last }
            : null;
    }

    /// <summary>
    /// Where the cells of an area go, for something every cell carries (data validation, a
    /// conditional format): an area across the band's edge is split, the part in the band moving
    /// and the parts beside it staying. Each part is given as it was and as it is, in the order of
    /// the area's cells; gone parts are left out.
    /// </summary>
    public List<(Area Before, Area After)> MapCells(Area area)
    {
        var (first, last) = Across(area);
        if (Band is not { } band || Covers(first, last) || Misses(first, last))
            return Map(area) is { } moved ? [(area, moved)] : [];
        var parts = new List<(Area, Area)>(3);
        if (first < band.First)
            parts.Add((WithAcross(area, first, band.First - 1), WithAcross(area, first, band.First - 1)));
        var inside = WithAcross(area, Math.Max(first, band.First), Math.Min(last, band.Last));
        if (Map(inside) is { } after)
            parts.Add((inside, after));
        if (last > band.Last)
            parts.Add((WithAcross(area, band.Last + 1, last), WithAcross(area, band.Last + 1, last)));
        return parts;
    }

    /// <summary>Whether an area holds cells that move or are deleted.</summary>
    public bool Touches(Area area)
    {
        var (first, last) = Across(area);
        if (Misses(first, last))
            return false;
        var along = Rows ? area.LastRow : area.LastColumn;
        return along >= At;
    }

    /// <summary>The area's span across the edit: its columns for rows, its rows for columns.</summary>
    public (int First, int Last) Across(Area area) => Rows ? (area.FirstColumn, area.LastColumn) : (area.FirstRow, area.LastRow);

    private Area WithAcross(Area area, int first, int last) =>
        Rows ? area with { FirstColumn = first, LastColumn = last } : area with { FirstRow = first, LastRow = last };
}

/// <summary>
/// Rewrites a formula for inserted or deleted rows or columns: references into the edited sheet
/// follow the cells they point at, ranges grow or shrink, references to deleted cells become
/// <c>#REF!</c>, and relative references are written again for the formula's new cell.
/// References through several sheets (<c>Sheet1:Sheet3!A1</c>), table references, names and text
/// inside INDIRECT do not change, as in Excel.
/// </summary>
internal static class ReferenceShift
{
    /// <returns>The formula with its references moved; the same instance when none changed.</returns>
    /// <param name="node">The formula.</param>
    /// <param name="formulaSheet">The sheet holding the formula: references without a sheet point into it.
    /// Null for a workbook-wide defined name, whose references without a sheet are left alone.</param>
    /// <param name="oldOrigin">The formula's cell before the edit.</param>
    /// <param name="newOrigin">The formula's cell after the edit.</param>
    /// <param name="edited">The sheet whose rows or columns changed.</param>
    /// <param name="edit">The change.</param>
    public static FormulaNode Apply(FormulaNode node, Worksheet? formulaSheet, CellAddress oldOrigin, CellAddress newOrigin, Worksheet edited, SheetEdit edit) =>
        new Shifter(formulaSheet, oldOrigin, newOrigin, edited, edit, null).Visit(node);

    /// <summary>The formula with references to tables that no longer exist turned into <c>#REF!</c>, as Excel does.</summary>
    /// <param name="node">The formula.</param>
    /// <param name="removed">The removed tables' names in upper case.</param>
    /// <param name="edited">The sheet the tables were on.</param>
    public static FormulaNode BreakTables(FormulaNode node, IReadOnlySet<string> removed, Worksheet edited) =>
        new Shifter(null, default, default, edited, null, removed).Visit(node);

    /// <summary>
    /// The formula copied from one cell to another, as pasting it in Excel copies it: relative
    /// references keep their offsets, and one that would point past the sheet's edge from the new
    /// cell becomes <c>#REF!</c> instead of wrapping around.
    /// </summary>
    /// <returns>The formula for the new cell; the same instance when no reference broke.</returns>
    /// <param name="node">The formula.</param>
    /// <param name="newOrigin">The cell the formula is copied to.</param>
    public static FormulaNode Copy(FormulaNode node, CellAddress newOrigin) =>
        new Shifter(null, newOrigin, newOrigin, null, null, null, copying: true).Visit(node);

    // With an edit, references move; with removed tables, references to them break.
    private sealed class Shifter(Worksheet? formulaSheet, CellAddress oldOrigin, CellAddress newOrigin, Worksheet? edited, SheetEdit? shift,
        IReadOnlySet<string>? removedTables, bool copying = false) : FormulaRewriter
    {
        private readonly SheetEdit edit = shift ?? default;

        protected override FormulaNode? Replace(FormulaNode node) => node switch
        {
            ReferenceNode reference when copying => Copied(reference),
            ReferenceNode reference when shift is not null => Reference(reference),
            StructuredReferenceNode { Reference.Table: { } table } when removedTables?.Contains(table.ToUpperInvariant()) == true => new ErrorNode(ErrorKind.Ref),
            _ => null,
        };

        // A copied reference stays as it is unless a relative part of it falls off the sheet.
        private FormulaNode Copied(ReferenceNode node)
        {
            var area = node.Area;
            var fits = area.Kind switch
            {
                AreaKind.Rows => Fits(area.First.Row, newOrigin.Row, CellAddress.MaxRow) && Fits(area.Last.Row, newOrigin.Row, CellAddress.MaxRow),
                AreaKind.Columns => Fits(area.First.Column, newOrigin.Column, CellAddress.MaxColumn) && Fits(area.Last.Column, newOrigin.Column, CellAddress.MaxColumn),
                _ => Fits(area.First.Row, newOrigin.Row, CellAddress.MaxRow) && Fits(area.First.Column, newOrigin.Column, CellAddress.MaxColumn)
                    && Fits(area.Last.Row, newOrigin.Row, CellAddress.MaxRow) && Fits(area.Last.Column, newOrigin.Column, CellAddress.MaxColumn),
            };
            if (fits)
                return node;
            return node.Sheet is null ? new ErrorNode(ErrorKind.Ref) : new RefErrorNode(node.Sheet);

            static bool Fits(AxisRef axis, int origin, int max) => axis.IsAbsolute || (origin + axis.Value >= 1 && origin + axis.Value <= max);
        }

        private FormulaNode Reference(ReferenceNode node)
        {
            var target = node.Sheet is null ? formulaSheet
                : node.Sheet.Last is null && edited!.Workbook.TryGetSheet(node.Sheet.First, out var named) ? named
                : null;
            var area = node.Area;
            if (target == edited)
            {
                var mapped = MapArea(area);
                if (mapped is null)
                    return node.Sheet is null ? new ErrorNode(ErrorKind.Ref) : new RefErrorNode(node.Sheet);
                area = mapped.Value;
            }

            area = Reencode(node.Area, area);
            return area == node.Area ? node : new ReferenceNode(node.Sheet, area);
        }

        // The area with its positions moved by the edit, still encoded for the old origin's
        // positions: AxisRefs hold absolute indexes here, re-encoded afterwards.
        private AreaRef? MapArea(AreaRef area)
        {
            switch (area.Kind)
            {
                case AreaKind.Cell:
                {
                    var row = area.First.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow);
                    var column = area.First.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn);
                    return edit.Map(new CellAddress(row, column)) is { } moved ? Absolute(area, moved.Row, moved.Column, moved.Row, moved.Column) : null;
                }

                case AreaKind.Range:
                {
                    var r1 = area.First.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow);
                    var c1 = area.First.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn);
                    var r2 = area.Last.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow);
                    var c2 = area.Last.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn);
                    // A range only some of the moved cells are in stays, as in Excel.
                    if (edit.Rows ? !edit.Covers(Math.Min(c1, c2), Math.Max(c1, c2)) : !edit.Covers(Math.Min(r1, r2), Math.Max(r1, r2)))
                        return Absolute(area, r1, c1, r2, c2);
                    if (edit.Rows)
                    {
                        if (edit.MapSpan(Math.Min(r1, r2), Math.Max(r1, r2), CellAddress.MaxRow) is not { } rows)
                            return null;
                        return r1 <= r2 ? Absolute(area, rows.First, c1, rows.Last, c2) : Absolute(area, rows.Last, c1, rows.First, c2);
                    }

                    if (edit.MapSpan(Math.Min(c1, c2), Math.Max(c1, c2), CellAddress.MaxColumn) is not { } columns)
                        return null;
                    return c1 <= c2 ? Absolute(area, r1, columns.First, r2, columns.Last) : Absolute(area, r1, columns.Last, r2, columns.First);
                }

                case AreaKind.Rows when edit.Rows && edit.Whole:
                {
                    var r1 = area.First.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow);
                    var r2 = area.Last.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow);
                    if (edit.MapSpan(Math.Min(r1, r2), Math.Max(r1, r2), CellAddress.MaxRow) is not { } rows)
                        return null;
                    return r1 <= r2 ? Absolute(area, rows.First, 0, rows.Last, 0) : Absolute(area, rows.Last, 0, rows.First, 0);
                }

                case AreaKind.Columns when !edit.Rows && edit.Whole:
                {
                    var c1 = area.First.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn);
                    var c2 = area.Last.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn);
                    if (edit.MapSpan(Math.Min(c1, c2), Math.Max(c1, c2), CellAddress.MaxColumn) is not { } columns)
                        return null;
                    return c1 <= c2 ? Absolute(area, 0, columns.First, 0, columns.Last) : Absolute(area, 0, columns.Last, 0, columns.First);
                }

                default:
                    // Whole columns do not change when rows move, nor whole rows when columns do, nor
                    // whole rows when only some of their cells move.
                    return Absolute(area,
                        area.Kind == AreaKind.Columns ? 0 : area.First.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow),
                        area.Kind == AreaKind.Rows ? 0 : area.First.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn),
                        area.Kind == AreaKind.Columns ? 0 : area.Last.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow),
                        area.Kind == AreaKind.Rows ? 0 : area.Last.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn));
            }
        }

        // An area with the given positions, every part marked absolute; Reencode restores each
        // part's own kind.
        private static AreaRef Absolute(AreaRef area, int r1, int c1, int r2, int c2) => new(area.Kind,
            new CellRef(AxisRef.Absolute(r1), AxisRef.Absolute(c1)), new CellRef(AxisRef.Absolute(r2), AxisRef.Absolute(c2)));

        // Positions back into the original's absolute and relative parts, relative ones counted from the new origin.
        private AreaRef Reencode(AreaRef original, AreaRef positions)
        {
            if (positions == original)
            {
                // Not edited: absolute positions are the old ones resolved.
                if (oldOrigin == newOrigin)
                    return original;
                positions = Absolute(original,
                    original.Kind == AreaKind.Columns ? 0 : original.First.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow),
                    original.Kind == AreaKind.Rows ? 0 : original.First.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn),
                    original.Kind == AreaKind.Columns ? 0 : original.Last.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow),
                    original.Kind == AreaKind.Rows ? 0 : original.Last.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn));
            }

            return new AreaRef(original.Kind,
                new CellRef(Axis(original.First.Row, positions.First.Row.Value, newOrigin.Row), Axis(original.First.Column, positions.First.Column.Value, newOrigin.Column)),
                new CellRef(Axis(original.Last.Row, positions.Last.Row.Value, newOrigin.Row), Axis(original.Last.Column, positions.Last.Column.Value, newOrigin.Column)));
        }

        // Parts a kind of area does not use (the column of a whole-row reference) stay as they were.
        private static AxisRef Axis(AxisRef original, int position, int origin) =>
            position == 0 ? original : original.IsAbsolute ? AxisRef.Absolute(position) : AxisRef.Relative(position - origin);
    }
}
