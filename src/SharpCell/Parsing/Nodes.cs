using System.Collections.Generic;

namespace SharpCell.Parsing;

/// <summary>
/// Immutable formula tree. References keep relative parts as offsets (see <see cref="AxisRef"/>),
/// so a tree does not depend on the cell it was parsed for. Explicit parentheses are kept as
/// <see cref="ParenthesesNode"/> so printing reproduces the formula as written.
/// </summary>
internal abstract record FormulaNode;

internal sealed record NumberNode(double Value) : FormulaNode;

internal sealed record TextNode(string Value) : FormulaNode;

internal sealed record BooleanNode(bool Value) : FormulaNode;

internal sealed record ErrorNode(ErrorKind Error) : FormulaNode;

/// <summary>An omitted argument: the middle of <c>IF(A1,,B1)</c>.</summary>
internal sealed record MissingNode : FormulaNode
{
    public static MissingNode Instance { get; } = new();
}

internal sealed record ReferenceNode(SheetPrefix? Sheet, AreaRef Area) : FormulaNode;

/// <summary><c>Sheet1!#REF!</c>: a reference whose target was deleted.</summary>
internal sealed record RefErrorNode(SheetPrefix Sheet) : FormulaNode;

internal sealed record NameNode(SheetPrefix? Sheet, string Name) : FormulaNode;

/// <summary>A table reference such as <c>Table1[Col]</c>, kept as raw text; not evaluated in v0.1.</summary>
internal sealed record StructuredReferenceNode(string Text) : FormulaNode;

internal enum UnaryOperator
{
    Negate,
    Plus,
    Percent,
}

internal sealed record UnaryNode(UnaryOperator Operator, FormulaNode Operand) : FormulaNode;

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

internal sealed record BinaryNode(BinaryOperator Operator, FormulaNode Left, FormulaNode Right) : FormulaNode;

/// <summary>A call by name. The name is upper-case with <c>_xlfn.</c>/<c>_xlws.</c> prefixes removed.</summary>
internal sealed record FunctionNode(string Name, IReadOnlyList<FormulaNode> Arguments) : FormulaNode;

/// <summary>A call of a function value: <c>LAMBDA(x,x+1)(2)</c>.</summary>
internal sealed record CallNode(FormulaNode Callee, IReadOnlyList<FormulaNode> Arguments) : FormulaNode;

internal sealed record ParenthesesNode(FormulaNode Inner) : FormulaNode;

/// <summary>An array literal <c>{1,2;3,4}</c>; elements are constants only.</summary>
internal sealed record ArrayNode(CellValue[,] Values) : FormulaNode;

/// <summary>The spill-range operator <c>A1#</c> (<c>_xlfn.ANCHORARRAY</c> in files).</summary>
internal sealed record SpillNode(FormulaNode Operand) : FormulaNode;

/// <summary>The implicit intersection operator <c>@</c> (<c>_xlfn.SINGLE</c> in files).</summary>
internal sealed record ImplicitIntersectionNode(FormulaNode Operand) : FormulaNode;
