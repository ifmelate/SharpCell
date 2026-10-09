using System;

namespace SharpCell;

/// <summary>Thrown when formula text cannot be parsed. This is the only exception the parser throws.</summary>
public sealed class FormulaParseException : Exception
{
    /// <summary>Creates the exception for a formula that cannot be parsed.</summary>
    /// <param name="message">What is wrong.</param>
    /// <param name="position">The zero-based character position where parsing failed.</param>
    public FormulaParseException(string message, int position)
        : base(message)
    {
        Position = position;
    }

    /// <summary>Zero-based index in the formula text where the problem was found.</summary>
    public int Position { get; }
}
