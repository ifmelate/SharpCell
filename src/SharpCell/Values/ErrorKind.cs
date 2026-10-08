using System;

namespace SharpCell;

/// <summary>Excel error values. Numeric values match the codes returned by <c>ERROR.TYPE</c>.</summary>
public enum ErrorKind
{
    Null = 1,
    Div0 = 2,
    Value = 3,
    Ref = 4,
    Name = 5,
    Num = 6,
    NA = 7,
    Spill = 9,
    Calc = 14,
}

/// <summary>Conversion between <see cref="ErrorKind"/> and its literal text (<c>#DIV/0!</c>).</summary>
public static class ErrorKinds
{
    private static readonly (ErrorKind Kind, string Text)[] Table =
    [
        (ErrorKind.Null, "#NULL!"),
        (ErrorKind.Div0, "#DIV/0!"),
        (ErrorKind.Value, "#VALUE!"),
        (ErrorKind.Ref, "#REF!"),
        (ErrorKind.Name, "#NAME?"),
        (ErrorKind.Num, "#NUM!"),
        (ErrorKind.NA, "#N/A"),
        (ErrorKind.Spill, "#SPILL!"),
        (ErrorKind.Calc, "#CALC!"),
    ];

    public static string ToText(this ErrorKind kind)
    {
        foreach (var (k, text) in Table)
        {
            if (k == kind)
                return text;
        }

        throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unknown error kind.");
    }

    /// <summary>Matches an error literal at the start of <paramref name="text"/>, ignoring case.</summary>
    internal static bool TryMatchPrefix(ReadOnlySpan<char> text, out ErrorKind kind, out int length)
    {
        foreach (var (k, literal) in Table)
        {
            if (text.StartsWith(literal, StringComparison.OrdinalIgnoreCase))
            {
                kind = k;
                length = literal.Length;
                return true;
            }
        }

        kind = default;
        length = 0;
        return false;
    }

    /// <summary>Parses an error literal, ignoring case. The whole span must be the literal.</summary>
    public static bool TryParse(ReadOnlySpan<char> text, out ErrorKind kind)
    {
        foreach (var (k, literal) in Table)
        {
            if (text.Equals(literal, StringComparison.OrdinalIgnoreCase))
            {
                kind = k;
                return true;
            }
        }

        kind = default;
        return false;
    }
}
