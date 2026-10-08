using SharpCell.Parsing;

namespace SharpCell;

/// <summary>Storage for one non-empty cell: a constant, or a formula with its last computed value.</summary>
internal sealed class CellData
{
    public CellValue Value;

    /// <summary>Formula text as set, always starting with '='; null for a constant.</summary>
    public string? FormulaText;

    public FormulaNode? Formula;
}
