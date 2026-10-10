using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using SharpCell.Evaluation;

namespace SharpCell.Xlsx;

/// <summary>
/// What a workbook read from a file keeps for saving: the file itself and the state it was read in,
/// so a save can write into the same file only what changed.
/// </summary>
internal sealed class XlsxSource(byte[] bytes, XlsxLimits limits, string workbookPart,
    IReadOnlyDictionary<Worksheet, SheetSource> sheets, WorkbookLook look) : IWorkbookSource
{
    public byte[] Bytes { get; } = bytes;

    public XlsxLimits Limits { get; } = limits;

    public string WorkbookPart { get; } = workbookPart;

    /// <summary>The sheets the file has, with their parts.</summary>
    public IReadOnlyDictionary<Worksheet, SheetSource> Sheets { get; } = sheets;

    /// <summary>Names, tables, settings and the default style as read.</summary>
    public WorkbookLook Look { get; } = look;

    /// <summary>The sheet's record; a sheet added since the workbook was read has an empty one and no part.</summary>
    public SheetSource SheetFor(Worksheet sheet) => Sheets.TryGetValue(sheet, out var source) ? source : SheetSource.New;

    /// <summary>The same file for a clone of the workbook: the bytes and the per-sheet records are shared.</summary>
    public IWorkbookSource CopyFor(Func<Worksheet, Worksheet> map)
    {
        var sheets = new Dictionary<Worksheet, SheetSource>(Sheets.Count);
        foreach (var (sheet, source) in Sheets)
            sheets.Add(map(sheet), source);
        return new XlsxSource(Bytes, Limits, WorkbookPart, sheets, Look);
    }

    public static XlsxSource Capture(byte[] bytes, XlsxLimits limits, string workbookPart, Workbook workbook,
        IReadOnlyDictionary<Worksheet, string> parts)
    {
        var sheets = new Dictionary<Worksheet, SheetSource>();
        foreach (var (sheet, part) in parts)
            sheets[sheet] = SheetSource.Capture(sheet, part);
        return new XlsxSource(bytes, limits, workbookPart, sheets, WorkbookLook.Of(workbook));
    }
}

/// <summary>A workbook's defined names, tables, settings and default style as read.</summary>
internal sealed class WorkbookLook
{
    private WorkbookLook(Workbook workbook)
    {
        DefaultStyle = workbook.DefaultStyle;
        DateSystem = workbook.DateSystem;
        Iteration = workbook.Iteration;
        foreach (var name in workbook.Names.All)
            Names[Key(name)] = name.Text;
        Tables = workbook.Tables.Select(t => t.Name.ToUpperInvariant()).ToHashSet(StringComparer.Ordinal);
    }

    public static WorkbookLook Of(Workbook workbook) => new(workbook);

    /// <summary>Cells without a style of their own showed it.</summary>
    public CellStyle DefaultStyle { get; }

    public DateSystem DateSystem { get; }

    public IterationSettings Iteration { get; }

    /// <summary>Each defined name's formula text, by scope (the sheet's name, or "") and upper-case name.</summary>
    public Dictionary<(string Scope, string Name), string> Names { get; } = [];

    /// <summary>The tables' names in upper case.</summary>
    public IReadOnlySet<string> Tables { get; }

    public static (string Scope, string Name) Key(NameDefinition name) => (name.Scope?.Name ?? "", name.Name);

    public bool NamesChanged(Workbook workbook) =>
        workbook.Names.All.Count() != Names.Count
        || workbook.Names.All.Any(name => !Names.TryGetValue(Key(name), out var text) || text != name.Text);
}

/// <summary>A worksheet's part in the file and its formulas as they were read.</summary>
internal sealed class SheetSource(string? part, IReadOnlyDictionary<CellAddress, LoadedFormula> formulas,
    IReadOnlyDictionary<CellAddress, string> formats, IReadOnlyDictionary<CellAddress, CellStyle> styles, SheetLook look)
{
    /// <summary>The record of a sheet the file does not have.</summary>
    public static SheetSource New { get; } = new(null, new Dictionary<CellAddress, LoadedFormula>(),
        new Dictionary<CellAddress, string>(), new Dictionary<CellAddress, CellStyle>(), SheetLook.Of(new Workbook().AddSheet("Sheet1")));

    public static SheetSource Capture(Worksheet sheet, string part)
    {
        var formulas = new Dictionary<CellAddress, LoadedFormula>();
        foreach (var cell in sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
        {
            if (cell.Data.FormulaText is { } text)
                formulas[new CellAddress(cell.Row, cell.Column)] = new LoadedFormula(text, cell.Data.IsLegacy, cell.Data.FixedArray);
        }

        var formats = new Dictionary<CellAddress, string>();
        foreach (var (address, format) in sheet.Formats)
            formats[address] = format.FormatCode;
        return new SheetSource(part, formulas, formats, new Dictionary<CellAddress, CellStyle>(sheet.Styles), SheetLook.Of(sheet));
    }

    /// <summary>The sheet's sizes, merges, panes and filter as read.</summary>
    public SheetLook Look { get; } = look;

    /// <summary>The worksheet part; null for a sheet added since the workbook was read.</summary>
    public string? Part { get; } = part;

    public IReadOnlyDictionary<CellAddress, LoadedFormula> Formulas { get; } = formulas;

    /// <summary>The number format codes other than General, as read.</summary>
    public IReadOnlyDictionary<CellAddress, string> Formats { get; } = formats;

    /// <summary>The cells' own styles, as read.</summary>
    public IReadOnlyDictionary<CellAddress, CellStyle> Styles { get; } = styles;
}

/// <summary>A formula cell as read: the formula a save leaves in the file must still be this one.</summary>
internal readonly record struct LoadedFormula(string Text, bool IsLegacy, Area? FixedArray);

/// <summary>Everything about a workbook other than its cells and their looks, as comparable text (for tests).</summary>
internal static class Structure
{
    public static string Of(Workbook workbook)
    {
        var text = new StringBuilder();
        text.Append("dates ").Append(workbook.DateSystem).Append('\n');
        text.Append("iteration ").Append(workbook.Iteration).Append('\n');
        foreach (var sheet in workbook.Sheets)
        {
            text.Append("sheet ").Append(sheet.Name).Append(" filter ").Append(sheet.FilterMode).Append(" hidden");
            foreach (var row in sheet.HiddenRows.Order())
                text.Append(' ').Append(row.ToString(CultureInfo.InvariantCulture));
            text.Append('\n');
        }

        foreach (var name in workbook.Names.All
                     .OrderBy(n => n.Scope?.Name, StringComparer.Ordinal).ThenBy(n => n.Name, StringComparer.Ordinal))
            text.Append("name ").Append(name.Scope?.Name).Append('!').Append(name.Name).Append(' ').Append(name.Text).Append('\n');

        foreach (var table in workbook.Tables)
        {
            text.Append("table ").Append(table.Worksheet.Name).Append('!').Append(table.Name).Append(' ').Append(table.Range)
                .Append(table.HasHeaderRow ? " header" : "").Append(table.HasTotalsRow ? " totals" : "")
                .Append(' ').Append(string.Join('|', table.Columns)).Append('\n');
        }

        return text.ToString();
    }
}

/// <summary>How the sheets look apart from cell styles, as comparable text (for tests).</summary>
internal static class Presentation
{
    public static string Of(Workbook workbook)
    {
        var text = new StringBuilder();
        foreach (var sheet in workbook.Sheets)
        {
            text.Append("sheet ").Append(sheet.Name)
                .Append(" column ").Append(Number(sheet.DefaultColumnWidth))
                .Append(" row ").Append(Number(sheet.DefaultRowHeight))
                .Append(" frozen ").Append(sheet.FrozenRows.ToString(CultureInfo.InvariantCulture)).Append(' ').Append(sheet.FrozenColumns.ToString(CultureInfo.InvariantCulture))
                .Append(" gridlines ").Append(sheet.ShowGridlines).Append('\n');
            foreach (var (column, width) in sheet.ColumnWidths.OrderBy(p => p.Key))
                text.Append("column ").Append(column.ToString(CultureInfo.InvariantCulture)).Append(' ').Append(Number(width)).Append('\n');
            foreach (var column in sheet.HiddenColumns.Order())
                text.Append("hidden column ").Append(column.ToString(CultureInfo.InvariantCulture)).Append('\n');
            foreach (var (row, height) in sheet.RowHeights.OrderBy(p => p.Key))
                text.Append("row ").Append(row.ToString(CultureInfo.InvariantCulture)).Append(' ').Append(Number(height)).Append('\n');
            foreach (var area in sheet.MergedAreaList)
                text.Append("merge ").Append(area.ToString()).Append('\n');
        }

        return text.ToString();
    }

    private static string Number(double? value) => value?.ToString("R", CultureInfo.InvariantCulture) ?? "-";
}
