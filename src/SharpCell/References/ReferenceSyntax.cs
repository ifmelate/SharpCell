using System;
using System.Globalization;
using System.Text;

namespace SharpCell;

/// <summary>Parsing and printing of sheet-less references in A1 and R1C1 notation.</summary>
internal static class ReferenceSyntax
{
    private enum PartKind
    {
        Cell,
        Row,
        Column,
    }

    /// <summary>Parses <c>A1</c>, <c>A1:B2</c>, <c>A:B</c> or <c>1:2</c>; the whole span must match.</summary>
    public static bool TryParseA1Area(ReadOnlySpan<char> text, CellAddress origin, out AreaRef area)
    {
        return TryParseArea(text, allowSingleLine: false, out area, (ReadOnlySpan<char> part, out PartKind kind, out CellRef cell) =>
            TryParseA1Part(part, origin, out kind, out cell));
    }

    /// <summary>Parses <c>R1C1</c>, <c>R[-1]C:RC[1]</c>, <c>R1:R2</c>, <c>C[1]</c> and similar.</summary>
    public static bool TryParseR1C1Area(ReadOnlySpan<char> text, out AreaRef area)
    {
        return TryParseArea(text, allowSingleLine: true, out area, TryParseR1C1Part);
    }

    public static string FormatA1(AreaRef area, CellAddress origin)
    {
        var sb = new StringBuilder();
        switch (area.Kind)
        {
            case AreaKind.Cell:
                AppendA1Cell(sb, area.First, origin);
                break;
            case AreaKind.Range:
                AppendA1Cell(sb, area.First, origin);
                sb.Append(':');
                AppendA1Cell(sb, area.Last, origin);
                break;
            case AreaKind.Rows:
                AppendA1Row(sb, area.First.Row, origin);
                sb.Append(':');
                AppendA1Row(sb, area.Last.Row, origin);
                break;
            case AreaKind.Columns:
                AppendA1Column(sb, area.First.Column, origin);
                sb.Append(':');
                AppendA1Column(sb, area.Last.Column, origin);
                break;
        }

        return sb.ToString();
    }

    public static string FormatR1C1(AreaRef area)
    {
        var sb = new StringBuilder();
        switch (area.Kind)
        {
            case AreaKind.Cell:
                AppendR1C1Axis(sb, 'R', area.First.Row);
                AppendR1C1Axis(sb, 'C', area.First.Column);
                break;
            case AreaKind.Range:
                AppendR1C1Axis(sb, 'R', area.First.Row);
                AppendR1C1Axis(sb, 'C', area.First.Column);
                sb.Append(':');
                AppendR1C1Axis(sb, 'R', area.Last.Row);
                AppendR1C1Axis(sb, 'C', area.Last.Column);
                break;
            case AreaKind.Rows:
                AppendR1C1Axis(sb, 'R', area.First.Row);
                if (area.Last.Row != area.First.Row)
                {
                    sb.Append(':');
                    AppendR1C1Axis(sb, 'R', area.Last.Row);
                }

                break;
            case AreaKind.Columns:
                AppendR1C1Axis(sb, 'C', area.First.Column);
                if (area.Last.Column != area.First.Column)
                {
                    sb.Append(':');
                    AppendR1C1Axis(sb, 'C', area.Last.Column);
                }

                break;
        }

        return sb.ToString();
    }

    private delegate bool PartParser(ReadOnlySpan<char> text, out PartKind kind, out CellRef cell);

    // A lone row or column ("R1", "C") is a whole-line reference in R1C1; in A1 it needs both
    // sides of a colon ("1:1", "A:A"), otherwise it is a number or a name.
    private static bool TryParseArea(ReadOnlySpan<char> text, bool allowSingleLine, out AreaRef area, PartParser parsePart)
    {
        area = default;
        var colon = text.IndexOf(':');
        if (colon < 0)
        {
            if (!parsePart(text, out var kind, out var cell) || (kind != PartKind.Cell && !allowSingleLine))
                return false;

            area = kind switch
            {
                PartKind.Cell => AreaRef.Cell(cell),
                PartKind.Row => AreaRef.Rows(cell.Row, cell.Row),
                _ => AreaRef.Columns(cell.Column, cell.Column),
            };
            return true;
        }

        if (!parsePart(text[..colon], out var firstKind, out var first)
            || !parsePart(text[(colon + 1)..], out var lastKind, out var last)
            || firstKind != lastKind)
            return false;

        area = firstKind switch
        {
            PartKind.Cell => AreaRef.Range(first, last),
            PartKind.Row => AreaRef.Rows(first.Row, last.Row),
            _ => AreaRef.Columns(first.Column, last.Column),
        };
        return true;
    }

    private static bool TryParseA1Part(ReadOnlySpan<char> text, CellAddress origin, out PartKind kind, out CellRef cell)
    {
        kind = default;
        cell = default;
        var i = 0;
        var firstDollar = TakeDollar(text, ref i);
        var lettersStart = i;
        while (i < text.Length && char.IsAsciiLetter(text[i]))
            i++;
        var letters = text[lettersStart..i];

        // In a row-only part ("$1") the leading dollar belongs to the row.
        var columnAbsolute = letters.Length > 0 && firstDollar;
        var rowAbsolute = letters.Length > 0 ? TakeDollar(text, ref i) : firstDollar;
        var digits = text[i..];

        AxisRef column = default;
        AxisRef row = default;
        if (letters.Length > 0)
        {
            if (!ColumnNames.TryParse(letters, out var index))
                return false;
            column = columnAbsolute ? AxisRef.Absolute(index) : AxisRef.Relative(index - origin.Column);
        }

        if (digits.Length > 0)
        {
            if (!RowNumbers.TryParse(digits, out var index))
                return false;
            row = rowAbsolute ? AxisRef.Absolute(index) : AxisRef.Relative(index - origin.Row);
        }
        else if (rowAbsolute || letters.Length == 0)
        {
            return false;
        }

        kind = letters.Length == 0 ? PartKind.Row : digits.Length == 0 ? PartKind.Column : PartKind.Cell;
        cell = new CellRef(row, column);
        return true;
    }

    private static bool TakeDollar(ReadOnlySpan<char> text, ref int i)
    {
        if (i < text.Length && text[i] == '$')
        {
            i++;
            return true;
        }

        return false;
    }

    private static bool TryParseR1C1Part(ReadOnlySpan<char> text, out PartKind kind, out CellRef cell)
    {
        kind = default;
        cell = default;
        var i = 0;
        AxisRef row = default;
        AxisRef column = default;
        var hasRow = false;
        var hasColumn = false;

        if (i < text.Length && (text[i] is 'R' or 'r'))
        {
            i++;
            if (!TryParseR1C1Axis(text, ref i, CellAddress.MaxRow, out row))
                return false;
            hasRow = true;
        }

        if (i < text.Length && (text[i] is 'C' or 'c'))
        {
            i++;
            if (!TryParseR1C1Axis(text, ref i, CellAddress.MaxColumn, out column))
                return false;
            hasColumn = true;
        }

        if (i != text.Length || (!hasRow && !hasColumn))
            return false;

        kind = hasRow && hasColumn ? PartKind.Cell : hasRow ? PartKind.Row : PartKind.Column;
        cell = new CellRef(row, column);
        return true;
    }

    // After the R or C letter: nothing (offset 0), digits (absolute) or [signed offset].
    private static bool TryParseR1C1Axis(ReadOnlySpan<char> text, ref int i, int max, out AxisRef axis)
    {
        axis = AxisRef.Relative(0);
        if (i >= text.Length)
            return true;

        if (text[i] == '[')
        {
            var close = text[i..].IndexOf(']');
            if (close < 0)
                return false;

            var inner = text.Slice(i + 1, close - 1);
            if (!int.TryParse(inner, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var offset)
                || Math.Abs((long)offset) >= max)
                return false;

            axis = AxisRef.Relative(offset);
            i += close + 1;
            return true;
        }

        var start = i;
        while (i < text.Length && char.IsAsciiDigit(text[i]))
            i++;
        if (i == start)
            return true;

        if (i - start > 7 || !int.TryParse(text[start..i], NumberStyles.None, CultureInfo.InvariantCulture, out var index)
            || index < 1 || index > max)
            return false;

        axis = AxisRef.Absolute(index);
        return true;
    }

    private static void AppendA1Cell(StringBuilder sb, CellRef cell, CellAddress origin)
    {
        AppendA1Column(sb, cell.Column, origin);
        AppendA1Row(sb, cell.Row, origin);
    }

    private static void AppendA1Column(StringBuilder sb, AxisRef column, CellAddress origin)
    {
        if (column.IsAbsolute)
            sb.Append('$');
        sb.Append(ColumnNames.ToName(column.Resolve(origin.Column, CellAddress.MaxColumn)));
    }

    private static void AppendA1Row(StringBuilder sb, AxisRef row, CellAddress origin)
    {
        if (row.IsAbsolute)
            sb.Append('$');
        sb.Append(row.Resolve(origin.Row, CellAddress.MaxRow).ToString(CultureInfo.InvariantCulture));
    }

    private static void AppendR1C1Axis(StringBuilder sb, char letter, AxisRef axis)
    {
        sb.Append(letter);
        if (axis.IsAbsolute)
            sb.Append(axis.Value.ToString(CultureInfo.InvariantCulture));
        else if (axis.Value != 0)
            sb.Append('[').Append(axis.Value.ToString(CultureInfo.InvariantCulture)).Append(']');
    }
}
