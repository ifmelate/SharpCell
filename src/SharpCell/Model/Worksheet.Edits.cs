using System;
using System.Collections.Generic;
using System.Linq;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell;

// Inserting and deleting rows and columns: cells move, and references to them follow, as in Excel.
public sealed partial class Worksheet
{
    private readonly List<SheetEdit> _edits = [];

    /// <summary>The rows and columns inserted and deleted since the sheet was made or read, in order.</summary>
    internal IReadOnlyList<SheetEdit> Edits => _edits;

    /// <summary>
    /// Inserts empty rows before <paramref name="row"/>, as Excel's Insert Sheet Rows does. Cells,
    /// styles, row heights, hidden rows, merged cells, tables and data validation below move down,
    /// and every reference to them follows: in formulas of every sheet, in defined names and in
    /// data validation. A range that spans the new rows grows; a table, merged area or validated
    /// range they are inserted into grows. Every formula is calculated again at the next
    /// <see cref="Workbook.Recalculate"/>.
    /// </summary>
    /// <param name="row">The row the new rows take the place of, from 1.</param>
    /// <param name="count">How many rows to insert.</param>
    /// <exception cref="ArgumentOutOfRangeException">The rows would not fit on the sheet.</exception>
    /// <exception cref="InvalidOperationException">Content would be pushed past the last row, or the rows
    /// would split an array formula (Ctrl+Shift+Enter).</exception>
    /// <exception cref="NotSupportedException">A formula on this sheet, or one naming it, cannot be read by
    /// SharpCell (such as a link to another workbook), so its references could not follow.</exception>
    public void InsertRows(int row, int count = 1) => Edit(new SheetEdit(Rows: true, row, count, Delete: false), nameof(row));

    /// <summary>
    /// Deletes rows, as Excel's Delete Sheet Rows does: cells below move up. A reference to a
    /// deleted cell becomes <c>#REF!</c>, a range that loses some of its rows shrinks, and a table
    /// whose rows are all deleted is removed. Every formula is calculated again at the next
    /// <see cref="Workbook.Recalculate"/>.
    /// </summary>
    /// <param name="row">The first row to delete, from 1.</param>
    /// <param name="count">How many rows to delete.</param>
    /// <exception cref="ArgumentOutOfRangeException">The rows are not all on the sheet.</exception>
    /// <exception cref="InvalidOperationException">The rows hold part of an array formula, or a table's header or
    /// totals row or all its data rows without the whole table.</exception>
    /// <exception cref="NotSupportedException">A formula on this sheet, or one naming it, cannot be read by SharpCell.</exception>
    public void DeleteRows(int row, int count = 1) => Edit(new SheetEdit(Rows: true, row, count, Delete: true), nameof(row));

    /// <summary>Inserts empty columns before <paramref name="column"/>; see <see cref="InsertRows"/>.</summary>
    /// <param name="column">The column the new columns take the place of, from 1 (A is 1).</param>
    /// <param name="count">How many columns to insert.</param>
    /// <exception cref="ArgumentOutOfRangeException">The columns would not fit on the sheet.</exception>
    /// <exception cref="InvalidOperationException">Content would be pushed past the last column, or the columns
    /// would split an array formula.</exception>
    /// <exception cref="NotSupportedException">The columns would be inserted inside a table, or a formula on this
    /// sheet, or one naming it, cannot be read by SharpCell.</exception>
    public void InsertColumns(int column, int count = 1) => Edit(new SheetEdit(Rows: false, column, count, Delete: false), nameof(column));

    /// <summary>Deletes columns; see <see cref="DeleteRows"/>.</summary>
    /// <param name="column">The first column to delete, from 1 (A is 1).</param>
    /// <param name="count">How many columns to delete.</param>
    /// <exception cref="ArgumentOutOfRangeException">The columns are not all on the sheet.</exception>
    /// <exception cref="InvalidOperationException">The columns hold part of an array formula.</exception>
    /// <exception cref="NotSupportedException">The columns are some but not all of a table's columns, or a formula
    /// on this sheet, or one naming it, cannot be read by SharpCell.</exception>
    public void DeleteColumns(int column, int count = 1) => Edit(new SheetEdit(Rows: false, column, count, Delete: true), nameof(column));

    /// <summary>Inserts or deletes the cells of an area, moving the cells beside it; see <see cref="CellRange.Insert"/>.</summary>
    internal void EditCells(Area area, bool rows, bool delete)
    {
        var (at, count, band) = rows
            ? (area.FirstRow, area.Rows, (area.FirstColumn, area.LastColumn))
            : (area.FirstColumn, area.Columns, (area.FirstRow, area.LastRow));
        var whole = rows ? area.FirstColumn == 1 && area.LastColumn == CellAddress.MaxColumn : area.FirstRow == 1 && area.LastRow == CellAddress.MaxRow;
        Edit(new SheetEdit(rows, at, count, delete, whole ? null : band), "range");
    }

    private void Edit(SheetEdit edit, string parameter)
    {
        var max = edit.Rows ? CellAddress.MaxRow : CellAddress.MaxColumn;
        if (edit.At < 1 || edit.At > max)
            throw new ArgumentOutOfRangeException(parameter, edit.At, "The position is outside the sheet.");
        if (edit.Count < 1 || edit.Count > max - edit.At + 1)
            throw new ArgumentOutOfRangeException("count", edit.Count, "The rows or columns must fit on the sheet.");
        Workbook.ThrowIfInCustomFunction();
        Check(edit, max);
        Workbook.Calculation.Rebuild(() => Apply(edit));
        _edits.Add(edit);
    }

    // Everything that would make the edit fail is found before anything changes.
    private void Check(SheetEdit edit, int max)
    {
        // A formula SharpCell cannot parse (a link to another workbook) could not follow the cells.
        foreach (var sheet in Workbook.Sheets)
        {
            foreach (var cell in sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
            {
                if (cell.Data.Formula is UnsupportedNode unsupported && (sheet == this || Mentions(unsupported.Text)))
                    throw new NotSupportedException($"{sheet.Name}!{new CellAddress(cell.Row, cell.Column)} has a formula SharpCell cannot read; its references would not follow the moved cells.");
            }
        }

        foreach (var name in Workbook.Names.All)
        {
            if (name.Formula is UnsupportedNode unsupported && (name.Scope == this || Mentions(unsupported.Text)))
                throw new NotSupportedException($"The name '{name.Spelling}' has a formula SharpCell cannot read; its references would not follow the moved cells.");
        }

        foreach (var sheet in Workbook.Sheets)
        {
            foreach (var rule in sheet.ValidationRules)
            {
                foreach (var formula in new[] { rule.F1, rule.F2 })
                {
                    if (formula is UnsupportedNode unsupported && (sheet == this || Mentions(unsupported.Text)))
                        throw new NotSupportedException($"A validation rule on {sheet.Name}!{new CellRange(sheet, rule.Areas[0]).Address} has a formula SharpCell cannot read; its references would not follow the moved cells.");
                }
            }
        }

        if (!edit.Delete)
        {
            var last = max - edit.Count;
            bool Pushed(int row, int column) => (edit.Rows ? row : column) > last && edit.InBand(edit.Rows ? column : row);
            var pushed = Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn)
                    .Any(cell => Pushed(cell.Row, cell.Column) && !(cell.Data.Formula is null && cell.Data.SpillAnchor is not null))
                || _styles.Keys.Any(a => Pushed(a.Row, a.Column))
                || _formats.Keys.Any(a => Pushed(a.Row, a.Column))
                || _hyperlinks.Keys.Any(a => Pushed(a.Row, a.Column))
                || _validations.Exists(rule => rule.Areas.Exists(a => Pushed(a.FirstRow, a.FirstColumn)));
            if (pushed)
                throw new InvalidOperationException($"Content in the last {(edit.Rows ? "rows" : "columns")} would be pushed off the sheet.");
        }

        foreach (var cell in Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
        {
            if (cell.Data.FixedArray is not { } array)
                continue;
            var (first, last) = edit.Rows ? (array.FirstRow, array.LastRow) : (array.FirstColumn, array.LastColumn);
            if (Splits(edit, first, last) || Shears(edit, array))
                throw new InvalidOperationException($"{Name}!{new CellAddress(cell.Row, cell.Column)} is an array formula; part of it cannot be inserted into or deleted.");
        }

        foreach (var merged in _merged)
        {
            if (Shears(edit, merged))
                throw new InvalidOperationException($"Cells of the merged area {new CellRange(this, merged).Address} would be moved apart; unmerge it first.");
        }

        foreach (var table in Workbook.Tables)
        {
            if (table.Worksheet != this)
                continue;
            var area = table.Area;
            if (Shears(edit, area))
                throw new InvalidOperationException($"Cells of table '{table.Name}' would be moved apart from the rest of it.");
            if (!edit.Covers(edit.Across(area).First, edit.Across(area).Last))
                continue;
            if (!edit.Rows)
            {
                if (Splits(edit, area.FirstColumn, area.LastColumn))
                    throw new NotSupportedException($"Columns cannot yet be inserted into or deleted from table '{table.Name}'.");
                continue;
            }

            if (!edit.Delete || !Overlaps(edit, area.FirstRow, area.LastRow) || Contains(edit, area.FirstRow, area.LastRow))
                continue;
            var end = edit.At + edit.Count - 1;
            if ((table.HasHeaderRow && Inside(area.FirstRow)) || (table.HasTotalsRow && Inside(area.LastRow))
                || (edit.At <= table.FirstDataRow && end >= table.LastDataRow))
                throw new InvalidOperationException($"Deleting these rows would leave table '{table.Name}' without its header, totals or data rows.");

            bool Inside(int row) => row >= edit.At && row <= end;
        }
    }

    private bool Mentions(string formula) => formula.Contains(Name, StringComparison.OrdinalIgnoreCase);

    // Inserting strictly inside a span, or deleting some but not all of it.
    private static bool Splits(SheetEdit edit, int first, int last) => edit.Delete
        ? Overlaps(edit, first, last) && !Contains(edit, first, last)
        : edit.At > first && edit.At <= last;

    // An area some of whose cells move while the rest stay: one across the edge of the band.
    private static bool Shears(SheetEdit edit, Area area) =>
        edit.Touches(area) && !edit.Covers(edit.Across(area).First, edit.Across(area).Last);

    private static bool Overlaps(SheetEdit edit, int first, int last) => edit.At <= last && edit.At + edit.Count - 1 >= first;

    private static bool Contains(SheetEdit edit, int first, int last) => edit.At <= first && edit.At + edit.Count - 1 >= last;

    private static CellAddress? Map(SheetEdit edit, int row, int column) => edit.Map(new CellAddress(row, column));

    private void Apply(SheetEdit edit)
    {
        var origin = new CellAddress(1, 1);
        foreach (var sheet in Workbook.Sheets)
        {
            foreach (var cell in sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
            {
                var data = cell.Data;
                if (data.Formula is null)
                    continue;
                var oldOrigin = new CellAddress(cell.Row, cell.Column);
                var newOrigin = sheet == this ? Map(edit, cell.Row, cell.Column) ?? oldOrigin : oldOrigin;
                var formula = ReferenceShift.Apply(data.Formula, sheet, oldOrigin, newOrigin, this, edit);
                if (formula != data.Formula || newOrigin != oldOrigin)
                {
                    data.Formula = formula;
                    if (formula is not UnsupportedNode)
                        data.FormulaText = "=" + FormulaPrinter.Print(formula, newOrigin);
                }

                if (sheet == this && data.FixedArray is { } array)
                    data.FixedArray = edit.Map(array);
            }
        }

        foreach (var name in Workbook.Names.All.ToList())
        {
            var formula = ReferenceShift.Apply(name.Formula, name.Scope, origin, origin, this, edit);
            if (formula != name.Formula)
                Workbook.Names.Set(new NameDefinition(name.Name, name.Spelling, "=" + FormulaPrinter.Print(formula, origin), formula, name.Scope));
        }

        foreach (var sheet in Workbook.Sheets)
            sheet.FollowEdit(this, edit);

        Store.Move((row, column) => Map(edit, row, column));
        MoveKeys(_styles, edit);
        MoveKeys(_formats, edit);
        MoveKeys(_hyperlinks, edit);
        // Rows and columns keep their heights and widths when only some of their cells move.
        if (edit.Whole && edit.Rows)
        {
            MoveIndexes(_rowHeights, edit, CellAddress.MaxRow);
            MoveIndexes(_hiddenRows, edit, CellAddress.MaxRow);
        }
        else if (edit.Whole)
        {
            MoveIndexes(_columnWidths, edit, CellAddress.MaxColumn);
            MoveIndexes(_hiddenColumns, edit, CellAddress.MaxColumn);
        }

        for (var i = _merged.Count - 1; i >= 0; i--)
        {
            if (edit.Map(_merged[i]) is { IsSingleCell: false } merged)
                _merged[i] = merged;
            else
                _merged.RemoveAt(i);
        }

        var before = Workbook.Tables.Where(table => table.Worksheet == this).Select(table => table.Name.ToUpperInvariant()).ToList();
        Workbook.MoveTables(this, table => edit.Map(table.Area));
        var removed = before.Where(name => !Workbook.TryGetTable(name, out _)).ToHashSet(StringComparer.Ordinal);
        if (removed.Count > 0)
            BreakTableReferences(removed);
    }

    // Formulas naming a table whose rows were all deleted get #REF!, as in Excel.
    private void BreakTableReferences(HashSet<string> removed)
    {
        var origin = new CellAddress(1, 1);
        foreach (var sheet in Workbook.Sheets)
        {
            foreach (var cell in sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
            {
                if (cell.Data.Formula is not { } formula || formula is UnsupportedNode)
                    continue;
                var broken = ReferenceShift.BreakTables(formula, removed, this);
                if (broken != formula)
                {
                    cell.Data.Formula = broken;
                    cell.Data.FormulaText = "=" + FormulaPrinter.Print(broken, new CellAddress(cell.Row, cell.Column));
                }
            }
        }

        foreach (var name in Workbook.Names.All.ToList())
        {
            var broken = ReferenceShift.BreakTables(name.Formula, removed, this);
            if (broken != name.Formula)
                Workbook.Names.Set(new NameDefinition(name.Name, name.Spelling, "=" + FormulaPrinter.Print(broken, origin), broken, name.Scope));
        }
    }

    private void MoveKeys<T>(Dictionary<CellAddress, T> map, SheetEdit edit)
    {
        var moved = new List<(CellAddress, T)>(map.Count);
        foreach (var (address, value) in map)
        {
            if (Map(edit, address.Row, address.Column) is { } to)
                moved.Add((to, value));
        }

        map.Clear();
        foreach (var (address, value) in moved)
            map[address] = value;
    }

    private static void MoveIndexes(Dictionary<int, double> map, SheetEdit edit, int max)
    {
        var moved = map.Select(pair => (Index: edit.Map(pair.Key, max), pair.Value)).Where(p => p.Index is not null).ToList();
        map.Clear();
        foreach (var (index, value) in moved)
            map[index!.Value] = value;
    }

    private static void MoveIndexes(HashSet<int> set, SheetEdit edit, int max)
    {
        var moved = set.Select(index => edit.Map(index, max)).Where(index => index is not null).ToList();
        set.Clear();
        foreach (var index in moved)
            set.Add(index!.Value);
    }
}
