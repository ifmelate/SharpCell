using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using SharpCell.Evaluation;

namespace SharpCell.Xlsx;

/// <summary>
/// What a workbook read from a file keeps for saving: the file itself and the state it was read in,
/// so a save can write new values into the same file and refuse changes it cannot write.
/// </summary>
internal sealed class XlsxSource(byte[] bytes, XlsxLimits limits, string workbookPart,
    IReadOnlyDictionary<Worksheet, SheetSource> sheets, string structure, string presentation) : IWorkbookSource
{
    public byte[] Bytes { get; } = bytes;

    public XlsxLimits Limits { get; } = limits;

    public string WorkbookPart { get; } = workbookPart;

    public IReadOnlyDictionary<Worksheet, SheetSource> Sheets { get; } = sheets;

    /// <summary>The workbook's structure as read, from <see cref="Xlsx.Structure.Of"/>.</summary>
    public string Structure { get; } = structure;

    /// <summary>The sheets' geometry and the default style as read, from <see cref="Xlsx.Presentation.Of"/>.</summary>
    public string Presentation { get; } = presentation;

    /// <summary>The same file for a clone of the workbook: the bytes and the per-sheet records are shared.</summary>
    public IWorkbookSource CopyFor(Func<Worksheet, Worksheet> map)
    {
        var sheets = new Dictionary<Worksheet, SheetSource>(Sheets.Count);
        foreach (var (sheet, source) in Sheets)
            sheets.Add(map(sheet), source);
        return new XlsxSource(Bytes, Limits, WorkbookPart, sheets, Structure, Presentation);
    }

    public static XlsxSource Capture(byte[] bytes, XlsxLimits limits, string workbookPart, Workbook workbook,
        IReadOnlyDictionary<Worksheet, string> parts)
    {
        var sheets = new Dictionary<Worksheet, SheetSource>();
        foreach (var (sheet, part) in parts)
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
            sheets[sheet] = new SheetSource(part, formulas, formats, new Dictionary<CellAddress, CellStyle>(sheet.Styles));
        }

        return new XlsxSource(bytes, limits, workbookPart, sheets, Xlsx.Structure.Of(workbook), Xlsx.Presentation.Of(workbook));
    }
}

/// <summary>A worksheet's part in the file and its formulas as they were read.</summary>
internal sealed class SheetSource(string part, IReadOnlyDictionary<CellAddress, LoadedFormula> formulas,
    IReadOnlyDictionary<CellAddress, string> formats, IReadOnlyDictionary<CellAddress, CellStyle> styles)
{
    public string Part { get; } = part;

    public IReadOnlyDictionary<CellAddress, LoadedFormula> Formulas { get; } = formulas;

    /// <summary>The number format codes other than General, as read; a save refuses any change to them.</summary>
    public IReadOnlyDictionary<CellAddress, string> Formats { get; } = formats;

    /// <summary>The cells' own styles, as read; a save refuses any change to them.</summary>
    public IReadOnlyDictionary<CellAddress, CellStyle> Styles { get; } = styles;
}

/// <summary>A formula cell as read: the formula a save leaves in the file must still be this one.</summary>
internal readonly record struct LoadedFormula(string Text, bool IsLegacy, Area? FixedArray);

/// <summary>Everything about a workbook other than its cells that a save cannot write, as comparable text.</summary>
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

/// <summary>How the sheets look apart from cell styles, as comparable text: a save copies it from the file and cannot write changes.</summary>
internal static class Presentation
{
    public static string Of(Workbook workbook)
    {
        var text = new StringBuilder();
        text.Append("default style ").Append(workbook.DefaultStyle).Append('\n');
        foreach (var sheet in workbook.Sheets)
        {
            text.Append("sheet ").Append(sheet.Name)
                .Append(" column ").Append(Number(sheet.DefaultColumnWidth))
                .Append(" row ").Append(Number(sheet.DefaultRowHeight)).Append('\n');
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
