using System.Collections.Generic;

namespace SharpCell;

/// <summary>What one <see cref="Workbook.Recalculate"/> did.</summary>
public sealed class RecalculationResult
{
    internal RecalculationResult(IReadOnlyList<Cell> changedCells) => ChangedCells = changedCells;

    /// <summary>
    /// The cells whose values calculation changed since the previous <see cref="Workbook.Recalculate"/>
    /// (or since the workbook was loaded or cloned): formula cells, and cells a spill filled, changed
    /// or left (those are now empty). Cells you changed yourself are not included, nor cells that
    /// changed and changed back. In the order of their sheets in the workbook, then by row and column.
    /// </summary>
    public IReadOnlyList<Cell> ChangedCells { get; }
}
