using System.Collections.Generic;
using System.Globalization;

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

    public CellValue ReadCell(Worksheet sheet, int row, int column) =>
        sheet.Store.Get(row, column)?.Value ?? CellValue.Empty;

    public CellValue ReadCell(Worksheet sheet, StoredCell cell) => cell.Data.Value;

    public bool TryEnterName(NameDefinition name)
    {
        if (_activeNames.Count >= MaxNameDepth || _activeNames.Contains(name))
            return false;
        _activeNames.Add(name);
        return true;
    }

    public void LeaveName() => _activeNames.RemoveAt(_activeNames.Count - 1);
}
