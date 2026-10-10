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
    private bool _fullCalcRequested;

    private WritePlan(XlsxSource source, HashSet<CellKey> kept, bool rewriteAllValues)
    {
        Source = source;
        _kept = kept;
        RewriteAllValues = rewriteAllValues;
    }

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
        foreach (var sheet in workbook.Sheets)
            CheckFormulas(sheet, source.Sheets[sheet]);
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
        return new WritePlan(source, kept, options.RewriteAllValues);
    }

    private static void CheckFormulas(Worksheet sheet, SheetSource source)
    {
        foreach (var (address, loaded) in source.Formulas)
        {
            var data = sheet.Store.Get(address.Row, address.Column);
            if (data is null || data.FormulaText != loaded.Text || data.IsLegacy != loaded.IsLegacy || data.FixedArray != loaded.FixedArray)
                throw new NotSupportedException($"The formula in {sheet.Name}!{address} changed since the workbook was read; saving writes changed values, not formulas.");
        }

        foreach (var cell in sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
        {
            var address = new CellAddress(cell.Row, cell.Column);
            if (cell.Data.FormulaText is not null && !source.Formulas.ContainsKey(address))
                throw new NotSupportedException($"{sheet.Name}!{address} has a formula it did not have when the workbook was read; saving writes changed values, not formulas.");
        }
    }
}
