using System;
using System.Collections.Generic;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell;

// Copying cells, as copying and pasting them in Excel does.
public sealed partial class CellRange
{
    /// <summary>
    /// Copies the range's cells to another range, as copying and pasting them in Excel does: values,
    /// formulas, number formats, styles, hyperlinks, notes and data validation. A formula's relative references keep their distance
    /// from its cell, so <c>=A1*2</c> in B1 becomes <c>=A2*2</c> in B2, and a reference that would
    /// point past the sheet's edge becomes <c>#REF!</c>. Empty cells of the range clear the cells
    /// they are copied to. When the destination is a whole number of times as tall and as wide as
    /// the range, the range repeats to fill it; otherwise it is copied once, with its top-left cell
    /// at the destination's. The cells are read before any is written, so the ranges may overlap.
    /// The destination can be in another workbook: formulas then refer to that workbook's sheets
    /// and names. A cell filled by a formula's spill is copied as its value, unless that formula is
    /// copied too and spills again. Data validation formulas keep their distance from each cell as
    /// cell formulas do. Merged cells, column widths and row heights are not copied.
    /// Every formula reading the cells written is calculated again at the next
    /// <see cref="Workbook.Recalculate"/>.
    /// </summary>
    /// <param name="destination">Where to copy the cells to.</param>
    /// <returns>The cells written to.</returns>
    /// <exception cref="ArgumentException">The copy would not fit on the sheet.</exception>
    /// <exception cref="InvalidOperationException">A cell written to is part of an array formula other than its top-left cell.</exception>
    /// <exception cref="NotSupportedException">The range holds part of an array formula (Ctrl+Shift+Enter), or a formula
    /// SharpCell cannot read (such as a link to another workbook), in a cell or a data validation rule,
    /// whose references could not be moved.</exception>
    public CellRange CopyTo(CellRange destination)
    {
        ArgumentNullException.ThrowIfNull(destination);
        Worksheet.Workbook.ThrowIfInCustomFunction();
        destination.Worksheet.Workbook.ThrowIfInCustomFunction();
        var target = destination.Worksheet;
        var to = destination.Area;
        var (down, across) = to.Rows % Area.Rows == 0 && to.Columns % Area.Columns == 0
            ? (to.Rows / Area.Rows, to.Columns / Area.Columns)
            : (1, 1);
        var lastRow = (long)to.FirstRow + (long)Area.Rows * down - 1;
        var lastColumn = (long)to.FirstColumn + (long)Area.Columns * across - 1;
        if (lastRow > CellAddress.MaxRow || lastColumn > CellAddress.MaxColumn)
            throw new ArgumentException($"{Address} copied to {destination.Address} would run past the sheet's edge.", nameof(destination));
        var written = new Area(to.FirstRow, to.FirstColumn, (int)lastRow, (int)lastColumn);

        // Everything copied is read first, so a destination overlapping the range gets the cells as they were.
        var calculation = Worksheet.Workbook.Calculation;
        var cells = new List<Copied>();
        var positions = new HashSet<(int Row, int Column)>();
        foreach (var cell in Worksheet.Store.Enumerate(FirstRow, FirstColumn, LastRow, LastColumn))
        {
            var data = cell.Data;
            if (data.FixedArray is not null || calculation.FixedArrayAt(Worksheet, cell.Row, cell.Column) is not null)
                throw new NotSupportedException($"{new Cell(Worksheet, cell.Row, cell.Column)} is part of an array formula, which cannot be copied yet.");
            if (data.Formula is UnsupportedNode)
                throw new NotSupportedException($"{new Cell(Worksheet, cell.Row, cell.Column)} has a formula SharpCell cannot read; its references could not be moved.");
            var (row, column) = (cell.Row - FirstRow, cell.Column - FirstColumn);
            positions.Add((row, column));

            // A formula copied along spills again into the copies of its cells.
            if (data.Formula is null && data.SpillAnchor is { } anchor && anchor.Sheet == Worksheet && Area.Contains(anchor.Row, anchor.Column))
                continue;
            cells.Add(new Copied(row, column, data.Value, data.Formula, data.IsLegacy));
        }

        var styles = Within(Worksheet.Styles, Area);
        var formats = Within(Worksheet.Formats, Area);
        var links = Within(Worksheet.Hyperlinks, Area);
        var notes = Within(Worksheet.NoteMap, Area);
        var validations = Worksheet.ValidationsIn(Area);

        var targetCalculation = target.Workbook.Calculation;
        foreach (var cell in target.Store.Enumerate(written.FirstRow, written.FirstColumn, written.LastRow, written.LastColumn))
        {
            if (targetCalculation.FixedArrayAt(target, cell.Row, cell.Column) is { } member)
                throw new InvalidOperationException($"{new Cell(target, cell.Row, cell.Column)} is part of the array formula at {member}; change the whole array there.");
        }

        var tiles = new List<Area>();
        for (var r = 0; r < down; r++)
        {
            for (var c = 0; c < across; c++)
            {
                var top = to.FirstRow + r * Area.Rows;
                var left = to.FirstColumn + c * Area.Columns;
                tiles.Add(new Area(top, left, top + Area.Rows - 1, left + Area.Columns - 1));
            }
        }

        var rules = target.ValidationsPasted(validations, tiles);
        foreach (var tile in tiles)
        {
            Paste(target, tile, cells, positions, styles, formats);
            Replace(target.Hyperlinks, links, tile, target.SetHyperlink);
            Replace(target.NoteMap, notes, tile, target.SetNote);
        }

        target.SetValidationRules(rules);
        return new CellRange(target, written);
    }

    private static void Paste(Worksheet target, Area tile, List<Copied> cells, HashSet<(int Row, int Column)> positions,
        Dictionary<(int Row, int Column), CellStyle> styles, Dictionary<(int Row, int Column), Functions.NumberFormat> formats)
    {
        // Cells left empty in the range clear theirs; a cell spilled into is left to its formula.
        var cleared = new List<(int Row, int Column)>();
        foreach (var cell in target.Store.Enumerate(tile.FirstRow, tile.FirstColumn, tile.LastRow, tile.LastColumn))
        {
            if (!positions.Contains((cell.Row - tile.FirstRow, cell.Column - tile.FirstColumn))
                && !(cell.Data.Formula is null && cell.Data.SpillAnchor is not null))
                cleared.Add((cell.Row, cell.Column));
        }

        foreach (var (row, column) in cleared)
            new Cell(target, row, column).Value = CellValue.Empty;

        foreach (var copied in cells)
        {
            var cell = new Cell(target, tile.FirstRow + copied.Row, tile.FirstColumn + copied.Column);
            if (copied.Formula is { } formula)
            {
                var origin = new CellAddress(cell.Row, cell.Column);
                var node = ReferenceShift.Copy(formula, origin);
                cell.SetFormula("=" + FormulaPrinter.Print(node, origin), node, copied.Legacy);
            }
            else
            {
                cell.Value = copied.Value;
            }
        }

        Replace(target.Styles, styles, tile, target.SetStyle);
        Replace(target.Formats, formats, tile, target.SetFormat);
    }

    // What the cells of a tile have of their own, as styles, becomes what the copied cells had.
    private static void Replace<T>(IReadOnlyDictionary<CellAddress, T> map, Dictionary<(int Row, int Column), T> copied, Area tile, Action<int, int, T?> set)
        where T : class
    {
        foreach (var address in Within(map, tile).Keys)
        {
            if (!copied.ContainsKey(address))
                set(tile.FirstRow + address.Row, tile.FirstColumn + address.Column, null);
        }

        foreach (var ((row, column), value) in copied)
            set(tile.FirstRow + row, tile.FirstColumn + column, value);
    }

    // The entries of a map by cell that lie in an area, by their offset from its top-left cell.
    // A small area is looked up cell by cell; a large one, such as whole columns, filters the map.
    private static Dictionary<(int Row, int Column), T> Within<T>(IReadOnlyDictionary<CellAddress, T> map, Area area)
    {
        var found = new Dictionary<(int Row, int Column), T>();
        if (map.Count == 0)
            return found;
        if (area.CellCount <= map.Count)
        {
            for (var row = area.FirstRow; row <= area.LastRow; row++)
            {
                for (var column = area.FirstColumn; column <= area.LastColumn; column++)
                {
                    if (map.TryGetValue(new CellAddress(row, column), out var value))
                        found[(row - area.FirstRow, column - area.FirstColumn)] = value;
                }
            }
        }
        else
        {
            foreach (var (address, value) in map)
            {
                if (area.Contains(address.Row, address.Column))
                    found[(address.Row - area.FirstRow, address.Column - area.FirstColumn)] = value;
            }
        }

        return found;
    }

    // A cell read for copying, at its offset from the range's top-left cell.
    private readonly record struct Copied(int Row, int Column, CellValue Value, FormulaNode? Formula, bool Legacy);
}
