using System;
using System.Globalization;
using System.Text;

namespace SharpCell;

/// <summary>Number ↔ text conversions with Excel's rules, parameterized by the workbook culture.</summary>
internal static class NumberText
{
    // Excel keeps 15 significant digits when it turns a number into text.
    private const int SignificantDigits = 15;

    // Below this decimal exponent the General format switches to scientific notation.
    private const int MinFixedExponentGeneral = -9;

    // Number literals in formula text stay in fixed notation much longer: Excel writes
    // 3+0.0000000000000001 in files. The exact cut-off is not known; -20 covers what was seen.
    private const int MinFixedExponentLiteral = -20;

    /// <summary>
    /// Parses text the way Excel coerces it to a number: surrounding spaces, sign, parentheses for
    /// negatives, currency symbol, digit grouping, exponent and a trailing percent sign.
    /// Dates and times are not recognized here.
    /// </summary>
    public static bool TryParse(string text, CultureInfo culture, out double value)
    {
        value = 0;
        var format = culture.NumberFormat;
        var s = text.AsSpan().Trim();
        var negative = false;
        var hasSign = false;

        if (s.Length >= 2 && s[0] == '(' && s[^1] == ')')
        {
            negative = true;
            hasSign = true;
            s = s[1..^1].Trim();
        }

        if (!hasSign && TryTakeSign(ref s, out negative))
            hasSign = true;

        // The invariant culture's generic sign is not what anyone types; workbooks without a culture
        // behave like English Excel, where "$5" is 5.
        var currency = (culture.Name.Length == 0 ? "$" : format.CurrencySymbol).AsSpan();
        if (currency.Length > 0 && s.StartsWith(currency, StringComparison.Ordinal))
        {
            s = s[currency.Length..].TrimStart();
            if (!hasSign && TryTakeSign(ref s, out negative))
                hasSign = true;
        }
        else if (currency.Length > 0 && s.EndsWith(currency, StringComparison.Ordinal))
        {
            s = s[..^currency.Length].TrimEnd();
        }

        var percent = false;
        if (s.Length > 0 && s[^1] == '%')
        {
            percent = true;
            s = s[..^1].TrimEnd();
        }

        var invariant = new StringBuilder(s.Length);
        if (!TryNormalizeMantissa(s, format, invariant))
            return false;

        if (!double.TryParse(invariant.ToString(), NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent,
                CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number))
            return false;

        if (percent)
            number /= 100;
        value = negative ? -number : number;
        return true;
    }

    /// <summary>Formats a number the way Excel's General format does when converting to text.</summary>
    public static string FormatGeneral(double number, CultureInfo culture)
    {
        if (number == 0)
            return "0";

        // Rounding to 15 significant digits may carry into the exponent (9.9999999999999999 -> 10),
        // so digits and exponent are read back from the rounded representation.
        var rounded = number.ToString("E" + (SignificantDigits - 1), CultureInfo.InvariantCulture);
        return Layout(rounded, culture.NumberFormat.NumberDecimalSeparator, MinFixedExponentGeneral);
    }

    /// <summary>
    /// Formats a number literal for canonical formula text. Uses the shortest digits that read back
    /// as the same double, so printing never changes a value.
    /// </summary>
    public static string FormatLiteral(double number) =>
        number == 0 ? "0" : Layout(number.ToString("R", CultureInfo.InvariantCulture), ".", MinFixedExponentLiteral);

    // Re-lays out an invariant number string ("-1.5E-07", "0.30000000000000004", "1E+20") in
    // Excel's style: fixed notation for exponents in [minFixedExponent, 15), scientific otherwise.
    private static string Layout(string invariant, string separator, int minFixedExponent)
    {
        var s = invariant.AsSpan();
        var negative = s[0] == '-';
        if (negative)
            s = s[1..];

        var exponent = 0;
        var ePos = s.IndexOfAny('E', 'e');
        if (ePos >= 0)
        {
            exponent = int.Parse(s[(ePos + 1)..], NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
            s = s[..ePos];
        }

        var point = s.IndexOf('.');
        var integerPart = point < 0 ? s : s[..point];
        var fractionPart = point < 0 ? ReadOnlySpan<char>.Empty : s[(point + 1)..];
        var all = string.Concat(integerPart, fractionPart);
        var leadingZeros = all.Length - all.TrimStart('0').Length;
        var digits = all[leadingZeros..].TrimEnd('0');
        exponent += integerPart.Length - leadingZeros - 1;

        var sb = new StringBuilder();
        if (negative)
            sb.Append('-');

        if (exponent >= SignificantDigits || exponent < minFixedExponent)
        {
            sb.Append(digits[0]);
            if (digits.Length > 1)
                sb.Append(separator).Append(digits, 1, digits.Length - 1);
            sb.Append('E').Append(exponent < 0 ? '-' : '+')
                .Append(Math.Abs(exponent).ToString("00", CultureInfo.InvariantCulture));
        }
        else if (exponent >= 0)
        {
            var integerDigits = exponent + 1;
            if (digits.Length <= integerDigits)
            {
                sb.Append(digits).Append('0', integerDigits - digits.Length);
            }
            else
            {
                sb.Append(digits, 0, integerDigits).Append(separator)
                    .Append(digits, integerDigits, digits.Length - integerDigits);
            }
        }
        else
        {
            sb.Append('0').Append(separator).Append('0', -exponent - 1).Append(digits);
        }

        return sb.ToString();
    }

    private static bool TryTakeSign(ref ReadOnlySpan<char> s, out bool negative)
    {
        negative = false;
        if (s.Length == 0 || (s[0] != '-' && s[0] != '+'))
            return false;

        negative = s[0] == '-';
        s = s[1..].TrimStart();
        return true;
    }

    // Rewrites "1,234.5e-3" (in the culture's separators) as invariant "1234.5e-3".
    // Grouping is validated: the first group has 1-3 digits, every following group exactly 3.
    private static bool TryNormalizeMantissa(ReadOnlySpan<char> s, NumberFormatInfo format, StringBuilder output)
    {
        var decimalSeparator = format.NumberDecimalSeparator;
        var groupSeparator = format.NumberGroupSeparator;
        var i = 0;
        var digitCount = 0;
        var groupLength = 0;
        var groups = 0;

        while (i < s.Length)
        {
            if (char.IsAsciiDigit(s[i]))
            {
                output.Append(s[i]);
                digitCount++;
                groupLength++;
                i++;
            }
            else if (IsGroupSeparatorAt(s, i, groupSeparator, out var length))
            {
                if ((groups == 0 && (groupLength is 0 or > 3)) || (groups > 0 && groupLength != 3))
                    return false;
                groups++;
                groupLength = 0;
                i += length;
            }
            else
            {
                break;
            }
        }

        if (groups > 0 && groupLength != 3)
            return false;

        if (s[i..].StartsWith(decimalSeparator, StringComparison.Ordinal))
        {
            output.Append('.');
            i += decimalSeparator.Length;
            while (i < s.Length && char.IsAsciiDigit(s[i]))
            {
                output.Append(s[i]);
                digitCount++;
                i++;
            }
        }

        if (digitCount == 0)
            return false;

        if (i < s.Length && (s[i] == 'e' || s[i] == 'E'))
        {
            output.Append('e');
            i++;
            if (i < s.Length && (s[i] == '+' || s[i] == '-'))
                output.Append(s[i++]);

            var exponentDigits = 0;
            while (i < s.Length && char.IsAsciiDigit(s[i]))
            {
                output.Append(s[i]);
                exponentDigits++;
                i++;
            }

            if (exponentDigits == 0)
                return false;
        }

        return i == s.Length;
    }

    // Cultures that group with a no-break space (ru-RU, fr-FR) also accept a plain space,
    // because that is what people type.
    private static bool IsGroupSeparatorAt(ReadOnlySpan<char> s, int i, string groupSeparator, out int length)
    {
        length = 0;
        if (groupSeparator.Length == 0)
            return false;

        if (s[i..].StartsWith(groupSeparator, StringComparison.Ordinal))
        {
            length = groupSeparator.Length;
            return true;
        }

        if (groupSeparator.Length == 1 && char.IsWhiteSpace(groupSeparator[0]) && s[i] is ' ' or ' ' or ' ')
        {
            length = 1;
            return true;
        }

        return false;
    }
}
