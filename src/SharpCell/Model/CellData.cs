using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell;

/// <summary>Storage for one non-empty cell: a constant, or a formula with its last computed value.</summary>
internal sealed class CellData
{
    public CellValue Value;

    /// <summary>Formula text as set, always starting with '='; null for a constant.</summary>
    public string? FormulaText;

    public FormulaNode? Formula;

    /// <summary>The formula must be (re)calculated; <see cref="Value"/> is stale.</summary>
    public bool IsDirty;

    /// <summary>Evaluation started and waits for dirty inputs; meeting it again means a loop.</summary>
    public bool InProgress;

    /// <summary>The cell whose evaluation asked for this one while <see cref="InProgress"/>.</summary>
    public CellKey? Requester;

    /// <summary>Dependencies registered in the graph from the last completed evaluation.</summary>
    public Dependencies? Registered;

    /// <summary>What the last (possibly unfinished) evaluation read; used when a loop is resolved.</summary>
    public Dependencies? LastAttempt;

    /// <summary>Whether the last (possibly unfinished) evaluation used a volatile function.</summary>
    public bool LastAttemptVolatile;

    /// <summary>
    /// A formula written by Excel before dynamic arrays (no array flag in the file): its result is
    /// reduced by implicit intersection instead of spilling.
    /// </summary>
    public bool IsLegacy;

    /// <summary>For a cell filled by another cell's array result: that anchor. Such a cell has no formula.</summary>
    public CellKey? SpillAnchor;

    /// <summary>For an anchor whose array result spilled: the area it covers, anchor included.</summary>
    public Area? SpillArea;

    /// <summary>
    /// For a cell whose result is an array: the rectangle it wants to fill, spilled or not. Content
    /// appearing or disappearing there makes it try again. Kept apart from <see cref="Registered"/>:
    /// filling the rectangle is not reading it.
    /// </summary>
    public Area? SpillWatch;
}
