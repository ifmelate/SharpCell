namespace SharpCell.Parsing;

/// <summary>Limits that keep the parser safe on untrusted input.</summary>
internal static class FormulaLimits
{
    /// <summary>Excel's own limit on formula length.</summary>
    public const int MaxLength = 8192;

    /// <summary>
    /// Maximum nesting of parentheses, function calls, arrays and unary operators.
    /// Excel allows 64 nested function levels; this leaves headroom without risking the stack.
    /// </summary>
    public const int MaxDepth = 100;

    /// <summary>
    /// Maximum <see cref="FormulaNode.Depth"/> of a parsed tree. Postfix chains (<c>1%%%</c>,
    /// <c>A1##</c>, <c>f(1)(2)(3)</c>) grow a tree without nesting the text, so they are capped here.
    /// </summary>
    public const int MaxTreeDepth = 256;
}
