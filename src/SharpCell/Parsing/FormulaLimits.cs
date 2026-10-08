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
}
