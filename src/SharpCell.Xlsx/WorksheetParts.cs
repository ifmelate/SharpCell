using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using SharpCell.Evaluation;

namespace SharpCell.Xlsx;

/// <summary>The children of a worksheet part (CT_Worksheet) in the order the schema requires.</summary>
internal static class WorksheetParts
{
    private static readonly string[] Order =
    [
        "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection",
        "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells",
        "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
        "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
        "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts",
        "extLst",
    ];

    /// <summary>The position of an element in the schema's order, or -1 for one the schema does not list.</summary>
    public static int Rank(string localName) => Array.IndexOf(Order, localName);

    public static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);

    public static string Format(Area area) => area.IsSingleCell
        ? new CellAddress(area.FirstRow, area.FirstColumn).ToString()
        : $"{new CellAddress(area.FirstRow, area.FirstColumn)}:{new CellAddress(area.LastRow, area.LastColumn)}";
}

/// <summary>
/// How a sheet looked when it was read, apart from its cells: what the writer compares with the
/// sheet now to know which parts of the worksheet to write again.
/// </summary>
internal sealed class SheetLook
{
    private SheetLook(Worksheet sheet)
    {
        ColumnWidths = new Dictionary<int, double>(sheet.ColumnWidths);
        HiddenColumns = new HashSet<int>(sheet.HiddenColumns);
        RowHeights = new Dictionary<int, double>(sheet.RowHeights);
        HiddenRows = new HashSet<int>(sheet.HiddenRows);
        Merged = [.. sheet.MergedAreaList];
        Hyperlinks = new Dictionary<CellAddress, Hyperlink>(sheet.Hyperlinks);
        Notes = new Dictionary<CellAddress, Note>(sheet.NoteMap);
        ConditionalFormats = [.. sheet.ConditionalRules.Select(rule => ((IReadOnlyList<Area>)[.. rule.Areas], rule.Format))];
        Orientation = sheet.PageOrientation;
        Margins = sheet.PageMargins;
        PrintAreas = [.. sheet.PrintAreas];
        TitleRows = sheet.TitleRows;
        TitleColumns = sheet.TitleColumns;
        RowBreaks = [.. sheet.RowPageBreaks];
        ColumnBreaks = [.. sheet.ColumnPageBreaks];
        Validations = [.. sheet.ValidationRules.Select(rule => ((IReadOnlyList<Area>)rule.Areas.ToArray(), rule.Validation))];
        Protection = sheet.Protection;
        ProtectionPassword = sheet.ProtectionPassword;
        FrozenRows = sheet.FrozenRows;
        FrozenColumns = sheet.FrozenColumns;
        ShowGridlines = sheet.ShowGridlines;
        DefaultColumnWidth = sheet.DefaultColumnWidth;
        DefaultRowHeight = sheet.DefaultRowHeight;
        FilterMode = sheet.FilterMode;
        TabColor = sheet.TabColor;
        Visibility = sheet.Visibility;
        FilterArea = sheet.AutoFilter?.Area;
        FilterColumns = sheet.AutoFilter is { } filter ? [.. filter.Columns.Select(pair => (pair.Key, pair.Value))] : [];
    }

    public static SheetLook Of(Worksheet sheet) => new(sheet);

    private SheetLook(SheetLook look, SheetEditMap map, Worksheet sheet)
    {
        ColumnWidths = Moved(look.ColumnWidths, column => map.MapColumn(sheet, column));
        HiddenColumns = look.HiddenColumns.Select(column => map.MapColumn(sheet, column)).OfType<int>().ToHashSet();
        RowHeights = Moved(look.RowHeights, row => map.MapRow(sheet, row));
        HiddenRows = look.HiddenRows.Select(row => map.MapRow(sheet, row)).OfType<int>().ToHashSet();
        Merged = [.. look.Merged.Select(area => map.Map(sheet, area)).OfType<Area>().Where(area => !area.IsSingleCell)];
        var hyperlinks = new Dictionary<CellAddress, Hyperlink>();
        foreach (var (address, link) in look.Hyperlinks)
        {
            if (map.Map(sheet, address) is { } to)
                hyperlinks[to] = link;
        }

        Hyperlinks = hyperlinks;
        var notes = new Dictionary<CellAddress, Note>();
        foreach (var (address, note) in look.Notes)
        {
            if (map.Map(sheet, address) is { } to)
                notes[to] = note;
        }

        Notes = notes;
        // Rules whose cells moved compare by their ranges; their formulas, written for the moved first
        // cell, then differ and the rules are written again from the model.
        ConditionalFormats = [.. look.ConditionalFormats.Select(rule => ((IReadOnlyList<Area>)[.. rule.Areas.SelectMany(area => SheetEditMap.MapCells(sheet, area))], rule.Format))
            .Where(rule => rule.Item1.Count > 0)];
        Orientation = look.Orientation;
        Margins = look.Margins;
        PrintAreas = [.. look.PrintAreas.Select(area => map.Map(sheet, area)).OfType<Area>()];
        TitleRows = look.TitleRows is { } rows ? map.MapSpan(sheet, rows: true, rows.First, rows.Last) : null;
        TitleColumns = look.TitleColumns is { } columns ? map.MapSpan(sheet, rows: false, columns.First, columns.Last) : null;
        RowBreaks = [.. look.RowBreaks.Select(row => map.MapRow(sheet, row)).OfType<int>().Where(row => row >= 2).Distinct().Order()];
        ColumnBreaks = [.. look.ColumnBreaks.Select(column => map.MapColumn(sheet, column)).OfType<int>().Where(column => column >= 2).Distinct().Order()];

        // Not moved: the rules of a sheet whose rows or columns moved are always written again, and
        // moving them here would mean moving their formulas as the model does.
        Validations = look.Validations;
        Protection = look.Protection;
        ProtectionPassword = look.ProtectionPassword;
        FrozenRows = look.FrozenRows;
        FrozenColumns = look.FrozenColumns;
        ShowGridlines = look.ShowGridlines;
        DefaultColumnWidth = look.DefaultColumnWidth;
        DefaultRowHeight = look.DefaultRowHeight;
        FilterMode = look.FilterMode;
        TabColor = look.TabColor;
        Visibility = look.Visibility;
        FilterArea = look.FilterArea is { } filtered ? map.Map(sheet, filtered) : null;
        FilterColumns = [.. look.FilterColumns.Select(pair => (Column: map.MapColumn(sheet, pair.Column), pair.Filter))
            .Where(pair => pair.Column is not null).Select(pair => (pair.Column!.Value, pair.Filter))];
    }

    /// <summary>The sheet as read, with its rows and columns where the sheet's edits moved them.</summary>
    public SheetLook Mapped(SheetEditMap map, Worksheet sheet) => map.HasEdits(sheet) ? new SheetLook(this, map, sheet) : this;

    private static Dictionary<int, double> Moved(IReadOnlyDictionary<int, double> sizes, Func<int, int?> map)
    {
        var moved = new Dictionary<int, double>();
        foreach (var (index, size) in sizes)
        {
            if (map(index) is { } to)
                moved[to] = size;
        }

        return moved;
    }

    public IReadOnlyDictionary<int, double> ColumnWidths { get; }

    public IReadOnlySet<int> HiddenColumns { get; }

    public IReadOnlyDictionary<int, double> RowHeights { get; }

    public IReadOnlySet<int> HiddenRows { get; }

    public IReadOnlyList<Area> Merged { get; }

    /// <summary>The links the model held, by cell; elements the model leaves out are not here.</summary>
    public IReadOnlyDictionary<CellAddress, Hyperlink> Hyperlinks { get; }

    /// <summary>The conditional formats in priority order with their ranges, as read.</summary>
    public IReadOnlyList<(IReadOnlyList<Area> Areas, ConditionalFormat Format)> ConditionalFormats { get; }

    public PageOrientation Orientation { get; }

    public PageMargins Margins { get; }

    public IReadOnlyList<Area> PrintAreas { get; }

    public (int First, int Last)? TitleRows { get; }

    public (int First, int Last)? TitleColumns { get; }

    public IReadOnlyList<int> RowBreaks { get; }

    public IReadOnlyList<int> ColumnBreaks { get; }

    /// <summary>The notes the model held, by cell; threaded comments are not here.</summary>
    public IReadOnlyDictionary<CellAddress, Note> Notes { get; }

    /// <summary>The data validation rules in order, with their ranges, as read; not moved by edits.</summary>
    public IReadOnlyList<(IReadOnlyList<Area> Areas, DataValidation Validation)> Validations { get; }

    /// <summary>The sheet's protection as read; null when it was not protected.</summary>
    public SheetProtection? Protection { get; }

    /// <summary>The protection's password as read; null for none.</summary>
    public ProtectionPassword? ProtectionPassword { get; }

    public int FrozenRows { get; }

    public int FrozenColumns { get; }

    public bool ShowGridlines { get; }

    public double? DefaultColumnWidth { get; }

    public double? DefaultRowHeight { get; }

    public bool FilterMode { get; }

    public CellColor? TabColor { get; }

    public SheetVisibility Visibility { get; }

    /// <summary>The filter's range, or null for none.</summary>
    public Area? FilterArea { get; }

    /// <summary>The filter's criteria by sheet column.</summary>
    public IReadOnlyList<(int Column, ColumnFilter Filter)> FilterColumns { get; }

    public bool AutoFilterChanged(Worksheet sheet) =>
        sheet.AutoFilter?.Area != FilterArea
        || !(sheet.AutoFilter?.Columns.Select(pair => (pair.Key, pair.Value)) ?? []).SequenceEqual(FilterColumns);

    public bool TabColorChanged(Worksheet sheet) => TabColor != sheet.TabColor;

    public bool ColumnsChanged(Worksheet sheet) => !Same(ColumnWidths, sheet.ColumnWidths) || !HiddenColumns.SetEquals(sheet.HiddenColumns);

    public bool RowsChanged(Worksheet sheet) => !Same(RowHeights, sheet.RowHeights) || !HiddenRows.SetEquals(sheet.HiddenRows);

    public bool RowChanged(Worksheet sheet, int row) =>
        LoadedHeight(row) != sheet.RowHeight(row) || HiddenRows.Contains(row) != sheet.IsRowHidden(row);

    public double? LoadedHeight(int row) => RowHeights.TryGetValue(row, out var height) ? height : null;

    public double? LoadedWidth(int column) => ColumnWidths.TryGetValue(column, out var width) ? width : null;

    public bool MergesChanged(Worksheet sheet) =>
        Merged.Count != sheet.MergedAreaList.Count || !Merged.ToHashSet().SetEquals(sheet.MergedAreaList);

    public bool HyperlinksChanged(Worksheet sheet) =>
        Hyperlinks.Count != sheet.Hyperlinks.Count
        || Hyperlinks.Any(pair => !sheet.Hyperlinks.TryGetValue(pair.Key, out var link) || !link.Equals(pair.Value));

    public bool ConditionalFormatsChanged(Worksheet sheet)
    {
        var rules = sheet.ConditionalRules;
        if (rules.Count != ConditionalFormats.Count)
            return true;
        for (var i = 0; i < rules.Count; i++)
        {
            if (!rules[i].Areas.SequenceEqual(ConditionalFormats[i].Areas) || !rules[i].Format.Equals(ConditionalFormats[i].Format))
                return true;
        }

        return false;
    }

    public bool OrientationChanged(Worksheet sheet) => Orientation != sheet.PageOrientation;

    public bool MarginsChanged(Worksheet sheet) => !Margins.Equals(sheet.PageMargins);

    public bool PrintAreaChanged(Worksheet sheet) => !PrintAreas.SequenceEqual(sheet.PrintAreas);

    public bool PrintTitlesChanged(Worksheet sheet) => TitleRows != sheet.TitleRows || TitleColumns != sheet.TitleColumns;

    public bool BreaksChanged(Worksheet sheet, bool rows) =>
        rows ? !RowBreaks.SequenceEqual(sheet.RowPageBreaks) : !ColumnBreaks.SequenceEqual(sheet.ColumnPageBreaks);

    public bool NotesChanged(Worksheet sheet) =>
        Notes.Count != sheet.NoteMap.Count
        || Notes.Any(pair => !sheet.NoteMap.TryGetValue(pair.Key, out var note) || !note.Equals(pair.Value));

    public bool ValidationsChanged(Worksheet sheet)
    {
        var rules = sheet.ValidationRules;
        if (rules.Count != Validations.Count)
            return true;
        for (var i = 0; i < rules.Count; i++)
        {
            if (!rules[i].Areas.SequenceEqual(Validations[i].Areas) || !rules[i].Validation.Equals(Validations[i].Validation))
                return true;
        }

        return false;
    }

    public bool ProtectionChanged(Worksheet sheet) =>
        !Equals(Protection, sheet.Protection) || !Equals(ProtectionPassword, sheet.ProtectionPassword);

    public bool PanesChanged(Worksheet sheet) => FrozenRows != sheet.FrozenRows || FrozenColumns != sheet.FrozenColumns;

    public bool ViewChanged(Worksheet sheet) => PanesChanged(sheet) || ShowGridlines != sheet.ShowGridlines;

    public bool DefaultsChanged(Worksheet sheet) => DefaultColumnWidth != sheet.DefaultColumnWidth || DefaultRowHeight != sheet.DefaultRowHeight;

    public bool FilterChanged(Worksheet sheet) => FilterMode != sheet.FilterMode;

    private static bool Same(IReadOnlyDictionary<int, double> a, IReadOnlyDictionary<int, double> b) =>
        a.Count == b.Count && a.All(pair => b.TryGetValue(pair.Key, out var value) && value == pair.Value);
}
