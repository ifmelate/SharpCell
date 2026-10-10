using System;
using System.Collections.Generic;
using SharpCell.Functions;
using SharpCell.Parsing;

namespace SharpCell.Evaluation;

/// <summary>
/// How a formula is saved. Excel reads a formula saved without the dynamic array flag the way
/// Excel did before dynamic arrays: a range or array where one value is expected is reduced by
/// implicit intersection. A formula SharpCell calculates as a dynamic one can be saved that way
/// only when such a reduction cannot happen; otherwise it is saved as a dynamic array formula.
/// </summary>
internal static class FormulaStorage
{
    // Functions whose result may be an array or a multi-cell range even when every value argument
    // is a single value.
    private static readonly HashSet<string> MayReturnArrays = new(StringComparer.Ordinal)
    {
        "INDEX", "OFFSET", "INDIRECT", "XLOOKUP", "LOOKUP", "TRANSPOSE", "MMULT", "MINVERSE", "MUNIT", "FREQUENCY",
        "TREND", "GROWTH", "LINEST", "LOGEST", "ROW", "COLUMN", "MODE.MULT", "LET", "LAMBDA", "MAP", "REDUCE", "SCAN",
        "BYROW", "BYCOL", "MAKEARRAY", "SEQUENCE", "RANDARRAY", "FILTER", "SORT", "SORTBY", "UNIQUE", "TAKE", "DROP",
        "EXPAND", "HSTACK", "VSTACK", "TOCOL", "TOROW", "WRAPCOLS", "WRAPROWS", "CHOOSECOLS", "CHOOSEROWS",
        "TEXTSPLIT", "REGEXEXTRACT", "GROUPBY", "PIVOTBY", "TRIMRANGE", "FILTERXML", "STOCKHISTORY",
    };

    // Functions that return one of their arguments: their result is one value when those are.
    private static readonly HashSet<string> Choosers = new(StringComparer.Ordinal)
    {
        "IF", "IFS", "IFERROR", "IFNA", "SWITCH", "CHOOSE",
    };

    private const int MaxNameDepth = 8;

    /// <summary>Whether the formula can be saved without the dynamic array flag and mean the same.</summary>
    /// <param name="node">The formula.</param>
    /// <param name="workbook">The workbook, for its functions and defined names.</param>
    /// <param name="sheet">The formula's sheet, for sheet-scoped names.</param>
    public static bool IsPlain(FormulaNode node, Workbook workbook, Worksheet sheet) =>
        !MayReduce(node, scalarExpected: true, workbook, sheet, 0);

    // Whether implicit intersection could change the value of the node in this position.
    private static bool MayReduce(FormulaNode node, bool scalarExpected, Workbook workbook, Worksheet sheet, int nameDepth)
    {
        switch (node)
        {
            case NumberNode or TextNode or BooleanNode or ErrorNode or MissingNode or RefErrorNode or UnsupportedNode:
                return false;
            case ReferenceNode reference:
                return scalarExpected && reference.Area.Kind != AreaKind.Cell;
            case NameNode name:
                return scalarExpected && !NameIsOneValue(name, workbook, sheet, nameDepth);
            case StructuredReferenceNode structured:
                return scalarExpected && !(structured.Reference is { Rows: TableRows.ThisRow, FirstColumn: { } first } r && r.LastColumn == first);
            case ArrayNode:
                return scalarExpected;
            case ParenthesesNode parentheses:
                return MayReduce(parentheses.Inner, scalarExpected, workbook, sheet, nameDepth);
            case UnaryNode unary:
                return MayReduce(unary.Operand, true, workbook, sheet, nameDepth);
            case BinaryNode { Operator: BinaryOperator.Range or BinaryOperator.Union or BinaryOperator.Intersect }:
                return scalarExpected;
            case BinaryNode binary:
                return BinaryMayReduce(binary, workbook, sheet, nameDepth);
            case FunctionNode function:
                return FunctionMayReduce(function, scalarExpected, workbook, sheet, nameDepth);
            default:
                // A lambda call, A1#, @ and anything new: saved with the flag, as Excel would.
                return true;
        }
    }

    // Long operator chains are walked along their left spine in a loop, as the printer does.
    private static bool BinaryMayReduce(BinaryNode node, Workbook workbook, Worksheet sheet, int nameDepth)
    {
        FormulaNode current = node;
        while (current is BinaryNode { Operator: not (BinaryOperator.Range or BinaryOperator.Union or BinaryOperator.Intersect) } binary)
        {
            if (MayReduce(binary.Right, true, workbook, sheet, nameDepth))
                return true;
            current = binary.Left;
        }

        return MayReduce(current, true, workbook, sheet, nameDepth);
    }

    private static bool FunctionMayReduce(FunctionNode function, bool scalarExpected, Workbook workbook, Worksheet sheet, int nameDepth)
    {
        if (!workbook.TryGetFunction(function.Name, out var info) || info is null || info.IsCustom || info.Status == FunctionStatus.NotImplemented)
            return true;
        if (scalarExpected && MayReturnArrays.Contains(function.Name) && !IsPositionOfOneCell(function))
            return true;

        var chooser = Choosers.Contains(function.Name);
        for (var i = 0; i < function.Arguments.Count; i++)
        {
            var kind = info.KindAt(i);
            if (kind == ArgumentKind.ArrayContext && !chooser)
                continue;
            var scalar = chooser || kind is ArgumentKind.Value or ArgumentKind.ScalarAny or ArgumentKind.Lazy;
            if (MayReduce(function.Arguments[i], scalar, workbook, sheet, nameDepth))
                return true;
        }

        return false;
    }

    // ROW() and COLUMN(A1) are one number.
    private static bool IsPositionOfOneCell(FunctionNode function) =>
        function.Name is "ROW" or "COLUMN"
        && (function.Arguments.Count == 0 || function.Arguments[0] is ReferenceNode { Area.Kind: AreaKind.Cell });

    // A name holding a constant or one cell is one value; anything else might not be.
    private static bool NameIsOneValue(NameNode name, Workbook workbook, Worksheet sheet, int nameDepth)
    {
        if (nameDepth >= MaxNameDepth)
            return false;
        var scope = name.Sheet is null ? sheet : null;
        var upper = name.Name.ToUpperInvariant();
        if (name.Sheet is not null && name.Sheet.Last is null && workbook.TryGetSheet(name.Sheet.First, out var prefixed))
            scope = prefixed;
        if (!(scope is not null && workbook.Names.TryGet(upper, scope, out var definition)) && !workbook.Names.TryGet(upper, null, out definition))
            return false;
        return !MayReduce(definition!.Formula, scalarExpected: true, workbook, sheet, nameDepth + 1);
    }
}
