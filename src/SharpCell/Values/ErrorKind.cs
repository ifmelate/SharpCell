using System;

namespace SharpCell;

/// <summary>Excel error values. Numeric values match the codes returned by <c>ERROR.TYPE</c>.</summary>
public enum ErrorKind
{
    /// <summary><c>#NULL!</c>: the intersection of two ranges is empty.</summary>
    Null = 1,
    /// <summary><c>#DIV/0!</c>: division by zero.</summary>
    Div0 = 2,
    /// <summary><c>#VALUE!</c>: an argument has the wrong type.</summary>
    Value = 3,
    /// <summary><c>#REF!</c>: a reference is not valid.</summary>
    Ref = 4,
    /// <summary><c>#NAME?</c>: an unknown function or name.</summary>
    Name = 5,
    /// <summary><c>#NUM!</c>: a number is out of range or a calculation does not converge.</summary>
    Num = 6,
    /// <summary><c>#N/A</c>: a value is not available, such as a lookup that found nothing.</summary>
    NA = 7,
    /// <summary><c>#SPILL!</c>: a dynamic array result cannot spill because cells are in the way.</summary>
    Spill = 9,
    /// <summary><c>#CALC!</c>: the calculation cannot produce a result, such as an empty array.</summary>
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

    /// <summary>The literal text of the error, such as <c>#DIV/0!</c>.</summary>
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
