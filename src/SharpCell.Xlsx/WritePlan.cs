using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Xlsx;

/// <summary>
/// Checks that a workbook can be saved into the file it was read from, and holds what the writers
/// share: the source, the formula cells that keep their saved result, and whether the file must ask
/// Excel to recalculate when it opens it.
/// </summary>
internal sealed class WritePlan
{
    private readonly HashSet<CellKey> _kept;
    private readonly HashSet<Worksheet> _formulasChanged;
    private bool _fullCalcRequested;

    private WritePlan(XlsxSource source, HashSet<CellKey> kept, HashSet<Worksheet> formulasChanged, bool rewriteAllValues)
    {
        Source = source;
        _kept = kept;
        _formulasChanged = formulasChanged;
        RewriteAllValues = rewriteAllValues;
    }

    /// <summary>Whether a formula of the sheet was added, removed or changed since the workbook was read.</summary>
    public bool FormulasChangedIn(Worksheet sheet) => _formulasChanged.Contains(sheet);

    public bool AnyFormulaChanged => _formulasChanged.Count > 0;

    public XlsxSource Source { get; }

    /// <summary>Formula cells SharpCell could not calculate: they keep the result saved in the file.</summary>
    public IReadOnlySet<CellKey> KeptCells => _kept;

    public bool RewriteAllValues { get; }

    /// <summary>Whether the file must ask Excel to recalculate every formula when it opens it.</summary>
    public bool FullCalcOnLoad => _kept.Count > 0 || _fullCalcRequested;

    public void RequestFullCalcOnLoad() => _fullCalcRequested = true;

    /// <exception cref="NotSupportedException">The workbook was not read from a file, or something other than values changed.</exception>
    /// <exception cref="InvalidOperationException">Some formulas are not calculated.</exception>
    /// <exception cref="XlsxWriteException">Some formulas cannot be calculated and the options do not keep them.</exception>
    public static WritePlan Create(Workbook workbook, XlsxWriteOptions options)
    {
        if (workbook.Source is not XlsxSource source)
            throw new NotSupportedException("Only a workbook read by XlsxReader can be saved: SharpCell writes new values into the file the workbook was read from.");
        if (Structure.Of(workbook) != source.Structure)
            throw new NotSupportedException("Sheets, defined names, tables, hidden rows, filters, the date system or the iteration settings changed since the workbook was read; only cell values can be saved.");
        if (Presentation.Of(workbook) != source.Presentation)
            throw new NotSupportedException("Column widths, hidden columns, row heights, merged cells, frozen panes, gridlines or the default style changed since the workbook was read; XlsxWriter writes values, not styles.");
        var formulasChanged = new HashSet<Worksheet>();
        foreach (var sheet in workbook.Sheets)
        {
            if (FormulasChanged(sheet, source.Sheets[sheet]))
                formulasChanged.Add(sheet);
            CheckFormats(sheet, source.Sheets[sheet]);
            CheckStyles(sheet, source.Sheets[sheet]);
        }
        if (workbook.Calculation.HasDirty)
            throw new InvalidOperationException("Some formulas are not calculated; call Workbook.Recalculate before saving.");

        var problems = new List<CalculationDiagnostic>();
        var kept = new HashSet<CellKey>();
        foreach (var diagnostic in workbook.Diagnostics)
        {
            if (diagnostic is not { Sheet: { } sheet, Address: { } address }
                || diagnostic.Kind is not (DiagnosticKind.UnsupportedFormula or DiagnosticKind.FunctionFailure or DiagnosticKind.LimitExceeded)
                || !CellAddress.TryParse(address, out var cell))
                continue;
            problems.Add(diagnostic);
            kept.Add(new CellKey(sheet, cell.Row, cell.Column));
        }

        if (problems.Count > 0 && !options.KeepUncalculated)
            throw new XlsxWriteException(problems);
        return new WritePlan(source, kept, formulasChanged, options.RewriteAllValues);
    }

    // Styles are copied as they were: a number format SharpCell changed cannot be written.
    private static void CheckFormats(Worksheet sheet, SheetSource source)
    {
        foreach (var (address, code) in source.Formats)
        {
            if (sheet.FormatAt(address.Row, address.Column)?.FormatCode != code)
                throw FormatChanged(sheet, address);
        }

        foreach (var (address, format) in sheet.Formats)
        {
            if (!source.Formats.ContainsKey(address))
                throw FormatChanged(sheet, address);
        }
    }

    // Styles are copied as they were: a style SharpCell changed cannot be written.
    private static void CheckStyles(Worksheet sheet, SheetSource source)
    {
        foreach (var (address, style) in source.Styles)
        {
            if (sheet.StyleAt(address.Row, address.Column) != style)
                throw StyleChanged(sheet, address);
        }

        foreach (var address in sheet.Styles.Keys)
        {
            if (!source.Styles.ContainsKey(address))
                throw StyleChanged(sheet, address);
        }
    }

    private static NotSupportedException StyleChanged(Worksheet sheet, CellAddress address) =>
        new($"The style of {sheet.Name}!{address} changed since the workbook was read; XlsxWriter writes values, not styles.");

    private static NotSupportedException FormatChanged(Worksheet sheet, CellAddress address) =>
        new($"The number format of {sheet.Name}!{address} changed since the workbook was read; XlsxWriter writes values, not formats.");

    /// <summary>Whether the formula a cell has now is the one it was read with.</summary>
    public static bool IsLoadedFormula(CellData? data, SheetSource source, CellAddress address) =>
        source.Formulas.TryGetValue(address, out var loaded)
            ? data is not null && data.FormulaText == loaded.Text && data.IsLegacy == loaded.IsLegacy && data.FixedArray == loaded.FixedArray
            : data?.FormulaText is null;

    private static bool FormulasChanged(Worksheet sheet, SheetSource source)
    {
        foreach (var address in source.Formulas.Keys)
        {
            if (!IsLoadedFormula(sheet.Store.Get(address.Row, address.Column), source, address))
                return true;
        }

        foreach (var cell in sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
        {
            if (cell.Data.FormulaText is not null && !source.Formulas.ContainsKey(new CellAddress(cell.Row, cell.Column)))
                return true;
        }

        return false;
    }
}
