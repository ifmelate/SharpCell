using System.Collections.Generic;
using System.Globalization;
using System.Threading;

namespace SharpCell.Evaluation;

/// <summary>Everything an evaluation needs besides the tree: workbook, current cell and settings.</summary>
internal sealed class EvaluationContext(Workbook workbook, Worksheet? sheet, CellAddress origin)
{
    // Names nest through other names; deeper chains than this are treated as circular.
    private const int MaxNameDepth = 64;

    private readonly List<NameDefinition> _activeNames = [];

    public Workbook Workbook { get; } = workbook;

    /// <summary>The sheet of the formula's cell; null for a formula evaluated outside any sheet.</summary>
    public Worksheet? Sheet { get; } = sheet;

    public CellAddress Origin { get; } = origin;

    public CultureInfo Culture => Workbook.Culture;

    public CancellationToken CancellationToken { get; init; }

    /// <summary>Set when a volatile function ran, so the cell is recalculated every time.</summary>
    public bool UsedVolatile { get; set; }

    // One entry per distinct message: an element-wise call can fail a million times the same way.
    public void Report(DiagnosticKind kind, string message)
    {
        foreach (var existing in _diagnostics)
        {
            if (existing.Kind == kind && existing.Message == message)
                return;
        }

        _diagnostics.Add(IsDetached || Sheet is null
            ? new CalculationDiagnostic(kind, null, null, message)
            : new CalculationDiagnostic(kind, Sheet, Origin.ToString(), message));
    }

    private readonly HashSet<CellKey> _pendingSet = [];
    private readonly List<CalculationDiagnostic> _diagnostics = [];

    /// <summary>What a read of a dirty cell returns; any value would do, the evaluation is discarded.</summary>
    public static CellValue PendingPlaceholder => CellValue.Error(ErrorKind.NA);

    /// <summary>LET names and LAMBDA parameters in effect; null outside any LET or LAMBDA.</summary>
    public Scope? Scope { get; set; }

    /// <summary>Lambda calls currently in progress.</summary>
    public int LambdaDepth { get; set; }

    /// <summary>The formula predates dynamic arrays; see <see cref="CellData.IsLegacy"/>.</summary>
    public bool Legacy { get; init; }

    /// <summary>A formula evaluated through <see cref="Workbook.Evaluate"/>, not in a cell.</summary>
    public bool IsDetached { get; init; }

    /// <summary>Problems reported by this evaluation; they count only if the evaluation is kept.</summary>
    public IReadOnlyList<CalculationDiagnostic> Diagnostics => _diagnostics;

    /// <summary>Records what the evaluation reads, for the dependency graph; null when not needed.</summary>
    public Dependencies? Dependencies { get; init; }

    /// <summary>
    /// Dirty formula cells the evaluation tried to read. When not empty the result is meaningless:
    /// these cells are computed and the formula is evaluated again.
    /// </summary>
    public List<CellKey> Pending { get; } = [];

    public CellValue ReadCell(Worksheet sheet, int row, int column) => Read(sheet, row, column, sheet.Store.Get(row, column));

    public CellValue ReadCell(Worksheet sheet, StoredCell cell) => Read(sheet, cell.Row, cell.Column, cell.Data);

    public void RecordReference(Reference reference)
    {
        if (Dependencies is null)
            return;
        foreach (var area in reference.Areas)
            Dependencies.Areas.Add(area);
    }

    public void RecordName(string upperName) => Dependencies?.Names.Add(upperName);

    private CellValue Read(Worksheet sheet, int row, int column, CellData? data)
    {
        if (data is null)
            return CellValue.Empty;
        if (data.IsDirty && data.Formula is not null)
            return Wait(new CellKey(sheet, row, column));

        // A spilled value is only as current as its anchor.
        if (data.SpillAnchor is { } anchor && anchor.Data is { IsDirty: true, Formula: not null })
            return Wait(anchor);

        return data.Value;
    }

    private CellValue Wait(CellKey key)
    {
        if (_pendingSet.Add(key))
            Pending.Add(key);
        return PendingPlaceholder;
    }

    public bool TryEnterName(NameDefinition name)
    {
        if (_activeNames.Count >= MaxNameDepth || _activeNames.Contains(name))
            return false;
        _activeNames.Add(name);
        return true;
    }

    public void LeaveName() => _activeNames.RemoveAt(_activeNames.Count - 1);
}
