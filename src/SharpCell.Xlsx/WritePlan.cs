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

    private WritePlan(Workbook workbook, XlsxSource source, HashSet<CellKey> kept, HashSet<Worksheet> formulasChanged, bool rewriteAllValues)
    {
        Workbook = workbook;
        Source = source;
        _kept = kept;
        _formulasChanged = formulasChanged;
        RewriteAllValues = rewriteAllValues;
    }

    /// <summary>Whether a formula of the sheet was added, removed or changed since the workbook was read.</summary>
    public bool FormulasChangedIn(Worksheet sheet) => _formulasChanged.Contains(sheet);

    public bool AnyFormulaChanged => _formulasChanged.Count > 0;

    public Workbook Workbook { get; }

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
        if (Structure.Of(workbook, withSheetState: false) != source.Structure)
            throw new NotSupportedException("Sheets, defined names, tables, the date system or the iteration settings changed since the workbook was read; they cannot be saved yet.");
        var formulasChanged = new HashSet<Worksheet>();
        foreach (var sheet in workbook.Sheets)
        {
            if (FormulasChanged(sheet, source.Sheets[sheet]))
                formulasChanged.Add(sheet);
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
        return new WritePlan(workbook, source, kept, formulasChanged, options.RewriteAllValues);
    }

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
