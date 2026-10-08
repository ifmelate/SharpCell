using System;
using System.Globalization;

namespace SharpCell;

/// <summary>An absolute cell position on a sheet. Row and column are 1-based, as in Excel.</summary>
internal readonly record struct CellAddress(int Row, int Column)
{
    public const int MaxRow = 1_048_576;
    public const int MaxColumn = 16_384;

    public static bool IsValid(int row, int column) => row is >= 1 and <= MaxRow && column is >= 1 and <= MaxColumn;

    /// <summary>Parses a plain address such as <c>B3</c> (no <c>$</c>), ignoring case.</summary>
    public static bool TryParse(ReadOnlySpan<char> text, out CellAddress address)
    {
        address = default;
        var letters = 0;
        while (letters < text.Length && char.IsAsciiLetter(text[letters]))
            letters++;

        if (!ColumnNames.TryParse(text[..letters], out var column)
            || !RowNumbers.TryParse(text[letters..], out var row))
            return false;

        address = new CellAddress(row, column);
        return true;
    }

    public override string ToString() => ColumnNames.ToName(Column) + Row.ToString(CultureInfo.InvariantCulture);
}

internal static class ColumnNames
{
    public static string ToName(int column)
    {
        if (column is < 1 or > CellAddress.MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(column), column, "Column is outside the sheet.");

        Span<char> buffer = stackalloc char[3];
        var position = buffer.Length;
        while (column > 0)
        {
            column--;
            buffer[--position] = (char)('A' + column % 26);
            column /= 26;
        }

        return new string(buffer[position..]);
    }

    /// <summary>Parses column letters (<c>A</c>..<c>XFD</c>), ignoring case.</summary>
    public static bool TryParse(ReadOnlySpan<char> text, out int column)
    {
        column = 0;
        if (text.Length is 0 or > 3)
            return false;

        foreach (var ch in text)
        {
            if (!char.IsAsciiLetter(ch))
                return false;
            column = column * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
        }

        return column <= CellAddress.MaxColumn;
    }
}

internal static class RowNumbers
{
    public static bool TryParse(ReadOnlySpan<char> text, out int row)
    {
        row = 0;
        if (text.Length is 0 or > 7)
            return false;

        foreach (var ch in text)
        {
            if (!char.IsAsciiDigit(ch))
                return false;
            row = row * 10 + (ch - '0');
        }

        return row is >= 1 and <= CellAddress.MaxRow;
    }
}
