using System;
using System.Collections.Generic;
using System.Linq;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Xlsx;

/// <summary>
/// Where the file's positions are in the model after rows and columns were inserted or deleted:
/// cells, ranges, lists of ranges (sqref) and formulas the model does not hold (conditional
/// formats, data validation, print areas). Those formulas also follow sheets renamed or removed.
/// </summary>
internal sealed class SheetEditMap
{
    private readonly List<Worksheet> _edited;

    // The file's sheets by name: those renamed since with their names now, those removed, and all in order.
    private readonly Dictionary<string, string> _renamed = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<string> _removed = [];
    private readonly List<string> _order = [];

    public SheetEditMap(Workbook workbook, XlsxSource source)
    {
        _edited = workbook.Sheets.Where(sheet => sheet.Edits.Count > 0).ToList();
        var live = source.LiveSheets().Where(p => p.Source.Part is not null).ToDictionary(p => p.Source.Part!, p => p.Sheet, StringComparer.OrdinalIgnoreCase);
        foreach (var file in source.FileSheets)
        {
            _order.Add(file.Name);
            if (file.Part is null)
                continue;
            if (!live.TryGetValue(file.Part, out var sheet))
                _removed.Add(file.Name);
            else if (sheet.Name != file.Name)
                _renamed[file.Name] = sheet.Name;
        }
    }

    /// <summary>Whether any sheet of the workbook has edits, or a sheet of the file was renamed or removed.</summary>
    public bool Any => _edited.Count > 0 || SheetsChanged;

    /// <summary>Whether a sheet of the file was renamed or removed.</summary>
    public bool SheetsChanged => _renamed.Count > 0 || _removed.Count > 0;

    /// <summary>The name a sheet of the file has now; null when it was removed.</summary>
    public string? SheetName(string fileName) =>
        _renamed.TryGetValue(fileName, out var name) ? name
        : _removed.Exists(removed => removed.Equals(fileName, StringComparison.OrdinalIgnoreCase)) ? null
        : fileName;

    public bool HasEdits(Worksheet sheet) => sheet.Edits.Count > 0;

    public bool HasColumnEdits(Worksheet sheet) => sheet.Edits.Any(edit => !edit.Rows && edit.Whole);

    /// <summary>Whether cells of the sheet were inserted or deleted with a shift, moving cells of a row to other rows.</summary>
    public bool HasCellEdits(Worksheet sheet) => sheet.Edits.Any(edit => !edit.Whole);

    /// <summary>Where a whole row goes, with its height and hidden flag; shifted cells do not move it.</summary>
    public int? MapRow(Worksheet sheet, int row)
    {
        int? current = row;
        foreach (var edit in sheet.Edits)
        {
            if (edit.Rows && edit.Whole && current is { } r)
                current = edit.Map(r, CellAddress.MaxRow);
        }

        return current;
    }

    /// <summary>Where a whole column goes, with its width and hidden flag; shifted cells do not move it.</summary>
    public int? MapColumn(Worksheet sheet, int column)
    {
        int? current = column;
        foreach (var edit in sheet.Edits)
        {
            if (!edit.Rows && edit.Whole && current is { } c)
                current = edit.Map(c, CellAddress.MaxColumn);
        }

        return current;
    }

    /// <summary>Where a span of whole rows (or columns) goes, as a reference to them does; null when it is gone.</summary>
    public (int First, int Last)? MapSpan(Worksheet sheet, bool rows, int first, int last)
    {
        (int First, int Last)? current = (first, last);
        foreach (var edit in sheet.Edits)
        {
            if (edit.Rows == rows && edit.Whole && current is { } span)
                current = edit.MapSpan(span.First, span.Last, rows ? CellAddress.MaxRow : CellAddress.MaxColumn);
        }

        return current;
    }

    public CellAddress? Map(Worksheet sheet, CellAddress address)
    {
        CellAddress? current = address;
        foreach (var edit in sheet.Edits)
        {
            if (current is { } cell)
                current = edit.Map(cell);
        }

        return current;
    }

    /// <summary>A range, grown, shrunk or moved as a reference to it would be; null when it is gone.</summary>
    public Area? Map(Worksheet sheet, Area area)
    {
        Area? current = area;
        foreach (var edit in sheet.Edits)
        {
            if (current is { } a)
                current = edit.Map(a);
        }

        return current;
    }

    /// <summary>A space-separated list of A1 ranges with each range moved; ranges that are gone are left out, null when none is left.</summary>
    public string? MapSqref(Worksheet sheet, string sqref)
    {
        var mapped = new List<string>();
        foreach (var part in sqref.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (!TryParse(part, out var area))
            {
                mapped.Add(part);
                continue;
            }

            foreach (var moved in MapCells(sheet, area))
                mapped.Add(WorksheetParts.Format(moved));
        }

        return mapped.Count == 0 ? null : string.Join(' ', mapped);
    }

    /// <summary>The cells of a range moved by each edit in turn: split where only some of them shift.</summary>
    public static List<Area> MapCells(Worksheet sheet, Area area)
    {
        var current = new List<Area> { area };
        foreach (var edit in sheet.Edits)
            current = current.SelectMany(part => edit.MapCells(part).Select(p => p.After)).ToList();
        return current;
    }

    /// <summary>The first cell of a sqref that the sheet's edits leave, in row and column order of its ranges.</summary>
    public CellAddress? FirstSurvivingCell(Worksheet sheet, string sqref)
    {
        var cells = HasCellEdits(sheet);
        foreach (var part in sqref.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (!TryParse(part, out var area))
                continue;
            int? row = null, column = null;
            for (var r = area.FirstRow; r <= area.LastRow && row is null; r++)
                row = MapRow(sheet, r) is null ? null : r;
            for (var c = area.FirstColumn; c <= area.LastColumn && column is null; c++)
                column = MapColumn(sheet, c) is null ? null : c;
            if (row is not { } first || column is not { } left)
                continue;
            if (!cells)
                return new CellAddress(first, left);

            // Shifted cells can be deleted anywhere in the range: look row by row, skipping deleted
            // rows; every cell looked at before the first one left is a deleted cell.
            for (var r = first; r <= area.LastRow; r++)
            {
                if (MapRow(sheet, r) is null)
                    continue;
                for (var c = left; c <= area.LastColumn; c++)
                {
                    if (MapColumn(sheet, c) is not null && Map(sheet, new CellAddress(r, c)) is not null)
                        return new CellAddress(r, c);
                }
            }
        }

        return null;
    }

    /// <summary>The first cell of the first range of a sqref, the origin of the formulas that go with it.</summary>
    public static CellAddress? FirstCell(string? sqref) =>
        sqref?.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() is { } first && TryParse(first, out var area)
            ? new CellAddress(area.FirstRow, area.FirstColumn)
            : null;

    /// <summary>
    /// A formula of the file that is not in a cell, with its references moved by every sheet's edits;
    /// the text as it was when it cannot be parsed or nothing in it moved.
    /// </summary>
    /// <param name="text">The formula in file syntax, without '='.</param>
    /// <param name="sheet">The sheet the formula belongs to, or null for a workbook-wide one.</param>
    /// <param name="origin">The cell its relative references count from, in the file.</param>
    /// <param name="anchor">The file's cell the formula is moved as if it were in, when the origin was
    /// deleted: the first cell of its range that is left, which becomes the range's first cell.</param>
    public string MapFormula(string text, Worksheet? sheet, CellAddress origin, CellAddress? anchor = null)
    {
        if (!Any || text.Length == 0)
            return text;
        FormulaNode node;
        try
        {
            node = FormulaParser.Parse("=" + text, origin);
        }
        catch (FormulaParseException)
        {
            return text;
        }

        var shifted = SheetsChanged ? FollowSheets(node) : node;
        var current = anchor ?? origin;
        foreach (var edited in _edited)
        {
            foreach (var edit in edited.Edits)
            {
                var moved = current;
                if (edited == sheet)
                    moved = edit.Map(current) ?? current;

                shifted = ReferenceShift.Apply(shifted, sheet, current, moved, edited, edit);
                current = moved;
            }
        }

        return shifted == node && current == origin ? text : FormulaPrinter.PrintForFile(shifted, current, _ => null);
    }

    // References into removed sheets break, as the model's did, then renamed sheets get their new names.
    private FormulaNode FollowSheets(FormulaNode node)
    {
        var order = new List<string>(_order);
        foreach (var removed in _removed)
        {
            node = SheetReferences.Remove(node, removed, order);
            order.RemoveAt(order.FindIndex(name => name.Equals(removed, StringComparison.OrdinalIgnoreCase)));
        }

        if (_renamed.Count == 0)
            return node;
        return SheetReferences.Map(node, prefix =>
        {
            var first = prefix.IsDeleted ? prefix.First : _renamed.GetValueOrDefault(prefix.First, prefix.First);
            var last = prefix.Last is null ? null : _renamed.GetValueOrDefault(prefix.Last, prefix.Last);
            return first == prefix.First && last == prefix.Last ? prefix : new SheetPrefix(first, last);
        });
    }

    private static bool TryParse(string text, out Area area)
    {
        area = default;
        if (!ReferenceSyntax.TryParseA1Area(text.Replace("$", "", StringComparison.Ordinal), new CellAddress(1, 1), out var parsed))
            return false;
        area = Area.Resolve(parsed, new CellAddress(1, 1));
        return true;
    }
}
