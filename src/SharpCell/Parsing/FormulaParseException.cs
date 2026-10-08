using System;

namespace SharpCell;

/// <summary>Thrown when formula text cannot be parsed. This is the only exception the parser throws.</summary>
public sealed class FormulaParseException : Exception
{
    public FormulaParseException(string message, int position)
        : base(message)
    {
        Position = position;
    }

    /// <summary>Zero-based index in the formula text where the problem was found.</summary>
    public int Position { get; }
}
