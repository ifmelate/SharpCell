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

    public void Report(DiagnosticKind kind, string message) =>
        Workbook.AddDiagnostic(new CalculationDiagnostic(kind, Sheet, Sheet is null ? null : Origin.ToString(), message));

    private readonly HashSet<CellKey> _pendingSet = [];

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
        {
            var key = new CellKey(sheet, row, column);
            if (_pendingSet.Add(key))
                Pending.Add(key);

            // Any value will do: the evaluation is discarded and repeated.
            return CellValue.Error(ErrorKind.NA);
        }

        return data.Value;
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
