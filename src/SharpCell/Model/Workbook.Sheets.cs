using System;
using System.Collections.Generic;
using System.Linq;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell;

// Removing, renaming and moving sheets: formulas follow, as in Excel.
public sealed partial class Workbook
{
    /// <summary>
    /// Removes a sheet, as Excel's Delete Sheet does. References to it become <c>#REF!</c>
    /// (<c>=Sheet2!A1</c> becomes <c>=#REF!A1</c>) in formulas of every other sheet, defined names, data
    /// validation and conditional formats; a reference through several sheets that starts or ends
    /// at it keeps the sheets left between its ends. Names local to the sheet and its tables go
    /// with it, and references to those tables become <c>#REF!</c>. Every formula is calculated
    /// again at the next <see cref="Recalculate"/>. The sheet object can still be read, but not
    /// changed (<see cref="Worksheet.IsRemoved"/>).
    /// </summary>
    /// <param name="sheet">A sheet of this workbook.</param>
    /// <exception cref="ArgumentException">The sheet is not in this workbook.</exception>
    /// <exception cref="InvalidOperationException">It is the last visible sheet while hidden ones remain: a workbook
    /// keeps at least one visible sheet.</exception>
    /// <exception cref="NotSupportedException">A formula naming the sheet cannot be read by SharpCell (such as a link
    /// to another workbook), so it could not be broken as Excel breaks it.</exception>
    public void RemoveSheet(Worksheet sheet)
    {
        CheckOwn(sheet);
        ThrowIfInCustomFunction();
        if (sheet.Visibility == SheetVisibility.Visible && _sheets.Count > 1 && VisibleSheets() == 1)
            throw new InvalidOperationException("A workbook must keep at least one visible sheet; show another sheet before removing this one.");
        CheckReadable(sheet, "broken");

        var order = _sheets.Select(s => s.Name).ToList();
        var tables = _tables.Where(table => table.Worksheet == sheet).Select(table => table.Name.ToUpperInvariant()).ToHashSet(StringComparer.Ordinal);
        Calculation.Rebuild(() =>
        {
            _sheets.Remove(sheet);
            Names.RemoveScope(sheet);
            RemoveTables(sheet);
            Rewrite(node =>
            {
                var removed = SheetReferences.Remove(node, sheet.Name, order);
                return tables.Count == 0 ? removed : ReferenceShift.BreakTables(removed, tables, sheet);
            });
        });
        sheet.IsRemoved = true;
    }

    /// <summary>
    /// Moves a sheet to another place in the order, as dragging its tab in Excel does. Formulas
    /// do not change, but a reference through several sheets (<c>Sheet1:Sheet3!A1</c>) takes in the
    /// sheets now between its ends, so every formula is calculated again at the next
    /// <see cref="Recalculate"/>.
    /// </summary>
    /// <param name="sheet">A sheet of this workbook.</param>
    /// <param name="index">Its new place in <see cref="Sheets"/>, from 0.</param>
    /// <exception cref="ArgumentException">The sheet is not in this workbook.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The index is not a place in <see cref="Sheets"/>.</exception>
    public void MoveSheet(Worksheet sheet, int index)
    {
        CheckOwn(sheet);
        if (index < 0 || index >= _sheets.Count)
            throw new ArgumentOutOfRangeException(nameof(index), index, "The index is not a place in the sheets.");
        ThrowIfInCustomFunction();
        var current = _sheets.IndexOf(sheet);
        if (current == index)
            return;
        Calculation.Rebuild(() =>
        {
            _sheets.RemoveAt(current);
            _sheets.Insert(index, sheet);
        });
    }

    /// <summary>Renames a sheet and every reference to it; see <see cref="Worksheet.Name"/>.</summary>
    internal void RenameSheet(Worksheet sheet, string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        sheet.ThrowIfCannotChange();
        ValidateSheetName(name);
        if (TryGetSheet(name, out var other) && other != sheet)
            throw new ArgumentException($"A sheet named '{name}' already exists.", nameof(name));
        if (name == sheet.Name)
            return;
        CheckReadable(sheet, "renamed");

        var old = sheet.Name;
        Calculation.Rebuild(() =>
        {
            sheet.SetName(name);
            Rewrite(node => SheetReferences.Rename(node, old, name));
        });
    }

    internal int VisibleSheets() => _sheets.Count(sheet => sheet.Visibility == SheetVisibility.Visible);

    private void CheckOwn(Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (sheet.Workbook != this || sheet.IsRemoved)
            throw new ArgumentException("The sheet is not in this workbook.", nameof(sheet));
    }

    // A formula SharpCell cannot parse (a link to another workbook) could not follow the sheet.
    private void CheckReadable(Worksheet sheet, string how)
    {
        bool Mentions(FormulaNode? node) => node is UnsupportedNode unsupported && unsupported.Text.Contains(sheet.Name, StringComparison.OrdinalIgnoreCase);
        foreach (var other in _sheets)
        {
            foreach (var cell in other.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
            {
                if (Mentions(cell.Data.Formula))
                    throw new NotSupportedException($"{other.Name}!{new CellAddress(cell.Row, cell.Column)} has a formula SharpCell cannot read; its references to sheet '{sheet.Name}' could not be {how}.");
            }

            if (other.ValidationRules.Any(rule => Mentions(rule.F1) || Mentions(rule.F2))
                || other.ConditionalRules.Any(rule => Mentions(rule.F1) || Mentions(rule.F2)))
                throw new NotSupportedException($"A rule on sheet '{other.Name}' has a formula SharpCell cannot read; its references to sheet '{sheet.Name}' could not be {how}.");
        }

        foreach (var name in Names.All)
        {
            if (Mentions(name.Formula))
                throw new NotSupportedException($"The name '{name.Spelling}' has a formula SharpCell cannot read; its references to sheet '{sheet.Name}' could not be {how}.");
        }
    }

    // Every formula of the workbook through the rewrite: cells, names, validation and conditional formats.
    private void Rewrite(Func<FormulaNode, FormulaNode> rewrite)
    {
        var origin = new CellAddress(1, 1);
        foreach (var sheet in _sheets)
        {
            foreach (var cell in sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
            {
                if (cell.Data.Formula is not { } formula || formula is UnsupportedNode)
                    continue;
                var rewritten = rewrite(formula);
                if (rewritten == formula)
                    continue;
                cell.Data.Formula = rewritten;
                cell.Data.FormulaText = "=" + FormulaPrinter.Print(rewritten, new CellAddress(cell.Row, cell.Column));
            }

            sheet.RewriteValidationFormulas(rewrite);
            sheet.RewriteConditionalFormulas(rewrite);
        }

        foreach (var name in Names.All.ToList())
        {
            var rewritten = rewrite(name.Formula);
            if (rewritten != name.Formula)
                Names.Set(new NameDefinition(name.Name, name.Spelling, "=" + FormulaPrinter.Print(rewritten, origin), rewritten, name.Scope));
        }
    }

    private void RemoveTables(Worksheet sheet)
    {
        foreach (var table in _tables.Where(table => table.Worksheet == sheet).ToList())
        {
            _tables.Remove(table);
            _tablesByName.Remove(table.Name.ToUpperInvariant());
            _tablesByCorner.Remove(new CellKey(sheet, table.Area.FirstRow, table.Area.FirstColumn));
        }

        _tableAreas.Remove(sheet);
    }
}
