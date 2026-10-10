using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Parsing;

/// <summary>
/// Rows or columns inserted into or deleted from a sheet: <paramref name="Count"/> of them at
/// <paramref name="At"/> (1-based).
/// </summary>
/// <param name="Rows">Rows; otherwise columns.</param>
/// <param name="At">The first row or column inserted or deleted.</param>
/// <param name="Count">How many.</param>
/// <param name="Delete">Deleted; otherwise inserted.</param>
internal readonly record struct SheetEdit(bool Rows, int At, int Count, bool Delete)
{
    /// <summary>Where a row or column goes; null when it is deleted or pushed past the sheet's end.</summary>
    public int? Map(int index, int max)
    {
        if (!Delete)
            return index < At ? index : index + Count <= max ? index + Count : null;
        return index < At ? index : index >= At + Count ? index - Count : null;
    }

    /// <summary>
    /// Where a span of rows or columns goes, as a range does: one inserted strictly inside it grows
    /// it, one at its start moves it; deleting part of it shrinks it. Null when all of it is deleted
    /// or pushed past the end.
    /// </summary>
    public (int First, int Last)? MapSpan(int first, int last, int max)
    {
        if (!Delete)
        {
            if (first >= At)
                return first + Count > max ? null : (first + Count, Math.Min(last + Count, max));
            return last >= At ? (first, Math.Min(last + Count, max)) : (first, last);
        }

        var end = At + Count - 1;
        if (first >= At && last <= end)
            return null;
        var newFirst = first < At ? first : first > end ? first - Count : At;
        var newLast = last < At ? last : last > end ? last - Count : At - 1;
        return (newFirst, newLast);
    }

    /// <summary>Where an area goes along the edited direction; null when it is gone.</summary>
    public Area? Map(Area area)
    {
        if (Rows)
        {
            return MapSpan(area.FirstRow, area.LastRow, CellAddress.MaxRow) is { } rows
                ? area with { FirstRow = rows.First, LastRow = rows.Last }
                : null;
        }

        return MapSpan(area.FirstColumn, area.LastColumn, CellAddress.MaxColumn) is { } columns
            ? area with { FirstColumn = columns.First, LastColumn = columns.Last }
            : null;
    }
}

/// <summary>
/// Rewrites a formula for inserted or deleted rows or columns: references into the edited sheet
/// follow the cells they point at, ranges grow or shrink, references to deleted cells become
/// <c>#REF!</c>, and relative references are written again for the formula's new cell.
/// References through several sheets (<c>Sheet1:Sheet3!A1</c>), table references, names and text
/// inside INDIRECT do not change, as in Excel.
/// </summary>
internal static class ReferenceShift
{
    /// <returns>The formula with its references moved; the same instance when none changed.</returns>
    /// <param name="node">The formula.</param>
    /// <param name="formulaSheet">The sheet holding the formula: references without a sheet point into it.</param>
    /// <param name="oldOrigin">The formula's cell before the edit.</param>
    /// <param name="newOrigin">The formula's cell after the edit.</param>
    /// <param name="edited">The sheet whose rows or columns changed.</param>
    /// <param name="edit">The change.</param>
    public static FormulaNode Apply(FormulaNode node, Worksheet formulaSheet, CellAddress oldOrigin, CellAddress newOrigin, Worksheet edited, SheetEdit edit) =>
        new Shifter(formulaSheet, oldOrigin, newOrigin, edited, edit).Visit(node);

    private sealed class Shifter(Worksheet formulaSheet, CellAddress oldOrigin, CellAddress newOrigin, Worksheet edited, SheetEdit edit)
    {
        public FormulaNode Visit(FormulaNode node)
        {
            switch (node)
            {
                case ReferenceNode reference:
                    return Reference(reference);
                case UnaryNode unary:
                    return Visit(unary.Operand) is var operand && operand != unary.Operand ? new UnaryNode(unary.Operator, operand) : node;
                case BinaryNode binary:
                    return Binary(binary);
                case ParenthesesNode parentheses:
                    return Visit(parentheses.Inner) is var inner && inner != parentheses.Inner ? new ParenthesesNode(inner) : node;
                case FunctionNode function:
                    return Arguments(function.Arguments) is { } changed ? new FunctionNode(function.Name, changed) : node;
                case CallNode call:
                {
                    var callee = Visit(call.Callee);
                    var arguments = Arguments(call.Arguments);
                    return callee != call.Callee || arguments is not null ? new CallNode(callee, arguments ?? call.Arguments) : node;
                }
                case SpillNode spill:
                    return Visit(spill.Operand) is var spilled && spilled != spill.Operand ? new SpillNode(spilled) : node;
                case ImplicitIntersectionNode intersection:
                    return Visit(intersection.Operand) is var operandNode && operandNode != intersection.Operand ? new ImplicitIntersectionNode(operandNode) : node;
                default:
                    return node;
            }
        }

        // Operator chains are walked along their left spine in a loop: a long A1+A2+...+An does not
        // recurse once per operator.
        private FormulaNode Binary(BinaryNode node)
        {
            var chain = new List<BinaryNode> { node };
            var current = node;
            while (current.Left is BinaryNode left)
            {
                chain.Add(left);
                current = left;
            }

            var result = Visit(current.Left);
            for (var i = chain.Count - 1; i >= 0; i--)
            {
                var binary = chain[i];
                var right = Visit(binary.Right);
                result = result == binary.Left && right == binary.Right ? binary : new BinaryNode(binary.Operator, result, right);
            }

            return result;
        }

        private List<FormulaNode>? Arguments(IReadOnlyList<FormulaNode> arguments)
        {
            List<FormulaNode>? changed = null;
            for (var i = 0; i < arguments.Count; i++)
            {
                var argument = Visit(arguments[i]);
                if (argument != arguments[i] && changed is null)
                {
                    changed = new List<FormulaNode>(arguments.Count);
                    for (var j = 0; j < i; j++)
                        changed.Add(arguments[j]);
                }

                changed?.Add(argument);
            }

            return changed;
        }

        private FormulaNode Reference(ReferenceNode node)
        {
            var target = node.Sheet is null ? formulaSheet
                : node.Sheet.Last is null && formulaSheet.Workbook.TryGetSheet(node.Sheet.First, out var named) ? named
                : null;
            var area = node.Area;
            if (target == edited)
            {
                var mapped = MapArea(area);
                if (mapped is null)
                    return node.Sheet is null ? new ErrorNode(ErrorKind.Ref) : new RefErrorNode(node.Sheet);
                area = mapped.Value;
            }

            area = Reencode(node.Area, area);
            return area == node.Area ? node : new ReferenceNode(node.Sheet, area);
        }

        // The area with its positions moved by the edit, still encoded for the old origin's
        // positions: AxisRefs hold absolute indexes here, re-encoded afterwards.
        private AreaRef? MapArea(AreaRef area)
        {
            switch (area.Kind)
            {
                case AreaKind.Cell:
                {
                    var row = area.First.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow);
                    var column = area.First.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn);
                    if (edit.Rows)
                        return edit.Map(row, CellAddress.MaxRow) is { } r ? Absolute(area, r, column, r, column) : null;
                    return edit.Map(column, CellAddress.MaxColumn) is { } c ? Absolute(area, row, c, row, c) : null;
                }

                case AreaKind.Range:
                {
                    var r1 = area.First.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow);
                    var c1 = area.First.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn);
                    var r2 = area.Last.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow);
                    var c2 = area.Last.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn);
                    if (edit.Rows)
                    {
                        if (edit.MapSpan(Math.Min(r1, r2), Math.Max(r1, r2), CellAddress.MaxRow) is not { } rows)
                            return null;
                        return r1 <= r2 ? Absolute(area, rows.First, c1, rows.Last, c2) : Absolute(area, rows.Last, c1, rows.First, c2);
                    }

                    if (edit.MapSpan(Math.Min(c1, c2), Math.Max(c1, c2), CellAddress.MaxColumn) is not { } columns)
                        return null;
                    return c1 <= c2 ? Absolute(area, r1, columns.First, r2, columns.Last) : Absolute(area, r1, columns.Last, r2, columns.First);
                }

                case AreaKind.Rows when edit.Rows:
                {
                    var r1 = area.First.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow);
                    var r2 = area.Last.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow);
                    if (edit.MapSpan(Math.Min(r1, r2), Math.Max(r1, r2), CellAddress.MaxRow) is not { } rows)
                        return null;
                    return r1 <= r2 ? Absolute(area, rows.First, 0, rows.Last, 0) : Absolute(area, rows.Last, 0, rows.First, 0);
                }

                case AreaKind.Columns when !edit.Rows:
                {
                    var c1 = area.First.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn);
                    var c2 = area.Last.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn);
                    if (edit.MapSpan(Math.Min(c1, c2), Math.Max(c1, c2), CellAddress.MaxColumn) is not { } columns)
                        return null;
                    return c1 <= c2 ? Absolute(area, 0, columns.First, 0, columns.Last) : Absolute(area, 0, columns.Last, 0, columns.First);
                }

                default:
                    // Whole columns do not change when rows move, nor whole rows when columns do.
                    return Absolute(area,
                        area.Kind == AreaKind.Columns ? 0 : area.First.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow),
                        area.Kind == AreaKind.Rows ? 0 : area.First.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn),
                        area.Kind == AreaKind.Columns ? 0 : area.Last.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow),
                        area.Kind == AreaKind.Rows ? 0 : area.Last.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn));
            }
        }

        // An area with the given positions, every part marked absolute; Reencode restores each
        // part's own kind.
        private static AreaRef Absolute(AreaRef area, int r1, int c1, int r2, int c2) => new(area.Kind,
            new CellRef(AxisRef.Absolute(r1), AxisRef.Absolute(c1)), new CellRef(AxisRef.Absolute(r2), AxisRef.Absolute(c2)));

        // Positions back into the original's absolute and relative parts, relative ones counted from the new origin.
        private AreaRef Reencode(AreaRef original, AreaRef positions)
        {
            if (positions == original)
            {
                // Not edited: absolute positions are the old ones resolved.
                if (oldOrigin == newOrigin)
                    return original;
                positions = Absolute(original,
                    original.Kind == AreaKind.Columns ? 0 : original.First.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow),
                    original.Kind == AreaKind.Rows ? 0 : original.First.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn),
                    original.Kind == AreaKind.Columns ? 0 : original.Last.Row.Resolve(oldOrigin.Row, CellAddress.MaxRow),
                    original.Kind == AreaKind.Rows ? 0 : original.Last.Column.Resolve(oldOrigin.Column, CellAddress.MaxColumn));
            }

            return new AreaRef(original.Kind,
                new CellRef(Axis(original.First.Row, positions.First.Row.Value, newOrigin.Row), Axis(original.First.Column, positions.First.Column.Value, newOrigin.Column)),
                new CellRef(Axis(original.Last.Row, positions.Last.Row.Value, newOrigin.Row), Axis(original.Last.Column, positions.Last.Column.Value, newOrigin.Column)));
        }

        // Parts a kind of area does not use (the column of a whole-row reference) stay as they were.
        private static AxisRef Axis(AxisRef original, int position, int origin) =>
            position == 0 ? original : original.IsAbsolute ? AxisRef.Absolute(position) : AxisRef.Relative(position - origin);
    }
}
