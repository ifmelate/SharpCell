using System;
using System.Collections.Generic;

namespace SharpCell.Parsing;

/// <summary>
/// Immutable formula tree. References keep relative parts as offsets (see <see cref="AxisRef"/>),
/// so a tree does not depend on the cell it was parsed for. Explicit parentheses are kept as
/// <see cref="ParenthesesNode"/> so printing reproduces the formula as written.
/// </summary>
/// <remarks>
/// Nodes are plain classes compared by reference: there is no generated deep Equals, GetHashCode
/// or ToString that could recurse through a tall tree.
/// <para>
/// <see cref="Depth"/> is the recursion depth a visitor needs when it walks the left spine of
/// binary operator chains in a loop (as <see cref="FormulaPrinter"/> does): left children that
/// are binary nodes add nothing, every other edge adds one. The parser rejects trees deeper than
/// <see cref="FormulaLimits.MaxTreeDepth"/>, so such a visitor cannot overflow the stack, while
/// long chains like <c>A1+A2+...+A2000</c> stay legal.
/// </para>
/// </remarks>
internal abstract class FormulaNode(int depth)
{
    public int Depth { get; } = depth;

    public override string ToString() => GetType().Name;

    protected static int MaxDepth(IReadOnlyList<FormulaNode> nodes)
    {
        var max = 0;
        foreach (var node in nodes)
            max = Math.Max(max, node.Depth);
        return max;
    }
}

internal sealed class NumberNode(double value) : FormulaNode(1)
{
    public double Value { get; } = value;
}

internal sealed class TextNode(string value) : FormulaNode(1)
{
    public string Value { get; } = value;
}

internal sealed class BooleanNode(bool value) : FormulaNode(1)
{
    public bool Value { get; } = value;
}

internal sealed class ErrorNode(ErrorKind error) : FormulaNode(1)
{
    public ErrorKind Error { get; } = error;
}

/// <summary>An omitted argument: the middle of <c>IF(A1,,B1)</c>.</summary>
internal sealed class MissingNode : FormulaNode
{
    private MissingNode()
        : base(1)
    {
    }

    public static MissingNode Instance { get; } = new();
}

internal sealed class ReferenceNode(SheetPrefix? sheet, AreaRef area) : FormulaNode(1)
{
    public SheetPrefix? Sheet { get; } = sheet;

    public AreaRef Area { get; } = area;
}

/// <summary><c>Sheet1!#REF!</c>: a reference whose target was deleted.</summary>
internal sealed class RefErrorNode(SheetPrefix sheet) : FormulaNode(1)
{
    public SheetPrefix Sheet { get; } = sheet;
}

internal sealed class NameNode(SheetPrefix? sheet, string name) : FormulaNode(1)
{
    public SheetPrefix? Sheet { get; } = sheet;

    public string Name { get; } = name;
}

/// <summary>A table reference such as <c>Table1[Col]</c>: the text as written and what it refers to.</summary>
internal sealed class StructuredReferenceNode(string text, StructuredReference reference) : FormulaNode(1)
{
    public string Text { get; } = text;

    public StructuredReference Reference { get; } = reference;
}

/// <summary>
/// A formula from a file that could not be parsed, kept as its raw text. It evaluates to
/// <c>#NAME?</c> and reports why, so one formula cannot stop a workbook from loading.
/// </summary>
internal sealed class UnsupportedNode(string text, string reason) : FormulaNode(1)
{
    /// <summary>The formula text without the leading '='.</summary>
    public string Text { get; } = text;

    public string Reason { get; } = reason;
}

internal enum UnaryOperator
{
    Negate,
    Plus,
    Percent,
}

internal sealed class UnaryNode(UnaryOperator op, FormulaNode operand) : FormulaNode(operand.Depth + 1)
{
    public UnaryOperator Operator { get; } = op;

    public FormulaNode Operand { get; } = operand;
}

internal enum BinaryOperator
{
    Range,
    Intersect,
    Union,
    Power,
    Multiply,
    Divide,
    Add,
    Subtract,
    Concat,
    Equal,
    NotEqual,
    Less,
    LessEqual,
    Greater,
    GreaterEqual,
}

internal sealed class BinaryNode(BinaryOperator op, FormulaNode left, FormulaNode right)
    : FormulaNode(Math.Max(left is BinaryNode ? left.Depth : left.Depth + 1, right.Depth + 1))
{
    public BinaryOperator Operator { get; } = op;

    public FormulaNode Left { get; } = left;

    public FormulaNode Right { get; } = right;
}

/// <summary>A call by name. The name is upper-case with <c>_xlfn.</c>/<c>_xlws.</c> prefixes removed.</summary>
internal sealed class FunctionNode(string name, IReadOnlyList<FormulaNode> arguments) : FormulaNode(MaxDepth(arguments) + 1)
{
    public string Name { get; } = name;

    public IReadOnlyList<FormulaNode> Arguments { get; } = arguments;
}

/// <summary>A call of a function value: <c>LAMBDA(x,x+1)(2)</c>.</summary>
internal sealed class CallNode(FormulaNode callee, IReadOnlyList<FormulaNode> arguments)
    : FormulaNode(Math.Max(callee.Depth, MaxDepth(arguments)) + 1)
{
    public FormulaNode Callee { get; } = callee;

    public IReadOnlyList<FormulaNode> Arguments { get; } = arguments;
}

internal sealed class ParenthesesNode(FormulaNode inner) : FormulaNode(inner.Depth + 1)
{
    public FormulaNode Inner { get; } = inner;
}

/// <summary>An array literal <c>{1,2;3,4}</c>; elements are constants only.</summary>
internal sealed class ArrayNode(CellValue[,] values) : FormulaNode(1)
{
    public CellValue[,] Values { get; } = values;
}

/// <summary>The spill-range operator <c>A1#</c> (<c>_xlfn.ANCHORARRAY</c> in files).</summary>
internal sealed class SpillNode(FormulaNode operand) : FormulaNode(operand.Depth + 1)
{
    public FormulaNode Operand { get; } = operand;
}

/// <summary>The implicit intersection operator <c>@</c> (<c>_xlfn.SINGLE</c> in files).</summary>
internal sealed class ImplicitIntersectionNode(FormulaNode operand) : FormulaNode(operand.Depth + 1)
{
    public FormulaNode Operand { get; } = operand;
}
