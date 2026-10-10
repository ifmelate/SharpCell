using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SharpCell.Functions;

/// <summary>
/// An Excel number format code, as TEXT reads it: <c>#,##0.00</c>, <c>0%</c>, <c>0.00E+00</c>,
/// <c># ?/?</c>, <c>dd/mm/yyyy h:mm AM/PM</c>, <c>[h]:mm:ss</c>, <c>"Total: "@</c>.
/// <para>
/// Codes are in English syntax whatever the culture: <c>.</c> is the decimal point, <c>,</c> the
/// thousands separator. The output uses the culture's decimal and group separators and its month
/// and day names. Up to four sections separated by <c>;</c> apply to positive numbers, negative
/// numbers, zero and text; <c>[&gt;100]</c>-style conditions replace the sign rules. Colours and
/// locale tags in brackets are ignored, <c>*</c> fills are left out (a function result has no
/// column width to fill), <c>_x</c> gives one space. Digits beyond Excel's 15 significant ones
/// show as zeros, and rounding is decimal (2.675 to two places is 2.68).
/// </para>
/// </summary>
internal sealed class NumberFormat
{
    // Excel shows at most 15 significant digits.
    private const int SignificantDigits = 15;

    // Parsed codes are reused: TEXT over a column parses the same code for every cell.
    private const int CacheLimit = 512;
    private static readonly ConcurrentDictionary<string, NumberFormat?> Cache = new(StringComparer.Ordinal);

    private readonly Section[] _numberSections;
    private readonly Section? _textSection;
    private readonly bool _hasConditions;

    private NumberFormat(Section[] numberSections, Section? textSection)
    {
        _numberSections = numberSections;
        _textSection = textSection;
        _hasConditions = Array.Exists(numberSections, s => s.Condition is not null);
    }

    private enum Code
    {
        Literal,
        Digit,
        Point,
        Comma,
        Percent,
        Exponent,
        Slash,
        At,
        General,
        Year,
        Month,
        Minute,
        Day,
        Hour,
        Second,
        ElapsedHours,
        ElapsedMinutes,
        ElapsedSeconds,
        SecondFraction,
        AmPm,
    }

    /// <summary>The code as it was parsed.</summary>
    public string FormatCode { get; private init; } = "";

    /// <summary>Parses a format code; false for codes Excel rejects (a digit placeholder among date codes, more than four sections).</summary>
    public static bool TryParse(string code, out NumberFormat? format)
    {
        if (!Cache.TryGetValue(code, out format))
        {
            format = Parse(code);
            if (Cache.Count >= CacheLimit)
                Cache.Clear();
            Cache[code] = format;
        }

        return format is not null;
    }

    /// <summary>
    /// Formats a number; null where Excel's TEXT gives <c>#VALUE!</c>: a date or time code applied
    /// to a negative number or one past 9999-12-31.
    /// </summary>
    public string? Format(double value, CultureInfo culture, DateSystem dateSystem)
    {
        var (section, shown) = Choose(value);
        if (section is null)
            return NumberText.FormatGeneral(value, culture);
        return section.Format(shown, culture, dateSystem);
    }

    /// <summary>Formats text: through the text section (its <c>@</c> stands for the text), unchanged without one.</summary>
    public string FormatText(string text) => _textSection is null ? text : _textSection.FormatText(text);

    /// <summary>
    /// Excel's General format as TEXT and a cell of default width show it: at most 11 characters
    /// besides the sign, switching to scientific notation (<c>1.23457E+11</c>) when the number
    /// does not fit.
    /// </summary>
    public static string FormatGeneral(double value, CultureInfo culture)
    {
        if (value == 0)
            return "0";

        var sb = new StringBuilder();
        if (value < 0)
            sb.Append('-');
        AppendGeneral(sb, Math.Abs(value), culture.NumberFormat.NumberDecimalSeparator);
        return sb.ToString();
    }

    // The section for a value and the value it shows: the negative section shows the magnitude.
    private (Section? Section, double Shown) Choose(double value)
    {
        var sections = _numberSections;
        if (_hasConditions)
        {
            var first = sections[0];
            var second = sections.Length > 1 ? sections[1] : null;
            if (first.Condition is { } c0 && c0.Matches(value))
                return (first, value);
            if (second?.Condition is { } c1 && c1.Matches(value))
                return (second, value);
            if (first.Condition is not null && second?.Condition is not null)
                return sections.Length > 2 ? (sections[2], value) : (null, value);
            if (first.Condition is null)
                return (first, value);
            return second is null ? (null, value) : (second, value);
        }

        if (sections.Length == 1 || value > 0 || (value == 0 && sections.Length == 2))
            return (sections[0], value);
        if (value < 0)
            return (sections[1], -value);
        return (sections[2], value);
    }

    private static NumberFormat? Parse(string code)
    {
        var parts = SplitSections(code);
        if (parts is null || parts.Count > 4)
            return null;

        var sections = new List<Section>(parts.Count);
        foreach (var part in parts)
        {
            var section = Section.Parse(part);
            if (section is null)
                return null;
            sections.Add(section);
        }

        // A fourth section is for text; so is a lone section with @ (and it shows numbers too).
        Section? text = null;
        if (sections.Count == 4)
        {
            text = sections[3];
            sections.RemoveAt(3);
        }
        else if (sections.Count == 1 && sections[0].HasAt)
        {
            text = sections[0];
        }

        return new NumberFormat([.. sections], text) { FormatCode = code };
    }

    // Splits on ';' outside quotes, escapes and brackets. Null for an unterminated quote or bracket.
    private static List<string>? SplitSections(string code)
    {
        var parts = new List<string>();
        var start = 0;
        for (var i = 0; i < code.Length; i++)
        {
            switch (code[i])
            {
                case '"':
                    var close = code.IndexOf('"', i + 1);
                    if (close < 0)
                        return null;
                    i = close;
                    break;
                case '[':
                    var end = code.IndexOf(']', i + 1);
                    if (end < 0)
                        return null;
                    i = end;
                    break;
                case '\\':
                case '_':
                case '*':
                    i++;
                    break;
                case ';':
                    parts.Add(code[start..i]);
                    start = i + 1;
                    break;
            }
        }

        parts.Add(code[start..]);
        return parts;
    }

    private readonly record struct Token(Code Code, string Text = "", int Length = 0)
    {
        public char Placeholder => Text[0];
    }

    private sealed class Condition(string op, double operand)
    {
        public bool Matches(double value) => op switch
        {
            "<" => value < operand,
            "<=" => value <= operand,
            ">" => value > operand,
            ">=" => value >= operand,
            "<>" => value != operand,
            _ => value == operand,
        };

        public static Condition? Parse(string text)
        {
            var length = text.StartsWith("<=", StringComparison.Ordinal) || text.StartsWith(">=", StringComparison.Ordinal)
                || text.StartsWith("<>", StringComparison.Ordinal) ? 2 : 1;
            return double.TryParse(text.AsSpan(length), NumberStyles.Float, CultureInfo.InvariantCulture, out var operand)
                ? new Condition(text[..length], operand)
                : null;
        }
    }

    private sealed class Section
    {
        private readonly List<Token> _tokens;
        private readonly bool _isDate;
        private readonly bool _hasDigits;
        private readonly bool _twelveHour;
        private readonly int _secondDigits;
        private readonly int _percent;
        private readonly int _scale;
        private readonly bool _grouping;
        private readonly int _decimals;
        private readonly int _exponent = -1;
        private readonly int _point = -1;
        private readonly int _slash = -1;

        private Section(List<Token> tokens, Condition? condition)
        {
            _tokens = tokens;
            Condition = condition;
            foreach (var token in tokens)
            {
                switch (token.Code)
                {
                    case Code.At:
                        HasAt = true;
                        break;
                    case Code.AmPm:
                        _twelveHour = true;
                        _isDate = true;
                        break;
                    case Code.SecondFraction:
                        _secondDigits = token.Length;
                        _isDate = true;
                        break;
                    case >= Code.Year:
                        _isDate = true;
                        break;
                }
            }

            if (_isDate)
                return;

            for (var i = 0; i < tokens.Count; i++)
            {
                switch (tokens[i].Code)
                {
                    case Code.Digit:
                        _hasDigits = true;
                        break;
                    case Code.Percent:
                        _percent++;
                        break;
                    case Code.Exponent when _exponent < 0:
                        _exponent = i;
                        break;
                    case Code.Point when _point < 0 && _exponent < 0:
                        _point = i;
                        break;
                    case Code.Slash when _slash < 0 && i > 0 && tokens[i - 1].Code == Code.Digit
                                         && i + 1 < tokens.Count && (tokens[i + 1].Code == Code.Digit || IsLiteralDigit(tokens[i + 1])):
                        _slash = i;
                        break;
                }
            }

            // Commas between digit placeholders group thousands; trailing ones divide by 1000 each.
            var mantissaEnd = _exponent >= 0 ? _exponent : tokens.Count;
            var integerEnd = _point >= 0 ? _point : mantissaEnd;
            var seenDigit = false;
            for (var i = 0; i < mantissaEnd; i++)
            {
                var token = tokens[i];
                if (token.Code == Code.Digit)
                {
                    seenDigit = true;
                    if (i > _point && _point >= 0)
                        _decimals++;
                    continue;
                }

                if (token.Code != Code.Comma)
                    continue;
                var regionEnd = i < integerEnd ? integerEnd : mantissaEnd;
                var digitFollows = false;
                for (var j = i + 1; j < regionEnd && !digitFollows; j++)
                    digitFollows = tokens[j].Code == Code.Digit;
                if (!seenDigit)
                    tokens[i] = new Token(Code.Literal, ",");
                else if (!digitFollows)
                    _scale++;
                else if (i < integerEnd)
                    _grouping = true;
            }
        }

        public Condition? Condition { get; }

        public bool HasAt { get; }

        public static Section? Parse(string text)
        {
            var tokens = new List<Token>();
            Condition? condition = null;
            for (var i = 0; i < text.Length; i++)
            {
                var c = text[i];
                switch (c)
                {
                    case '"':
                        var close = text.IndexOf('"', i + 1);
                        tokens.Add(new Token(Code.Literal, text[(i + 1)..close]));
                        i = close;
                        break;
                    case '\\':
                        if (i + 1 < text.Length)
                            tokens.Add(new Token(Code.Literal, text[++i].ToString()));
                        break;
                    case '_':
                        tokens.Add(new Token(Code.Literal, " "));
                        i++;
                        break;
                    case '*':
                        i++;
                        break;
                    case '[':
                        var end = text.IndexOf(']', i + 1);
                        var content = text[(i + 1)..end];
                        i = end;
                        if (content.Length > 0 && content[0] is '<' or '>' or '=')
                        {
                            condition = Condition.Parse(content);
                            if (condition is null)
                                return null;
                        }
                        else if (content.StartsWith('$'))
                        {
                            var dash = content.IndexOf('-');
                            var symbol = dash < 0 ? content[1..] : content[1..dash];
                            if (symbol.Length > 0)
                                tokens.Add(new Token(Code.Literal, symbol));
                        }
                        else if (content.Length > 0 && IsRun(content, out var letter))
                        {
                            var kind = letter switch
                            {
                                'h' => Code.ElapsedHours,
                                'm' => Code.ElapsedMinutes,
                                's' => Code.ElapsedSeconds,
                                _ => Code.Literal,
                            };
                            if (kind != Code.Literal)
                                tokens.Add(new Token(kind, Length: content.Length));
                        }

                        // Anything else in brackets is a colour or a locale tag.
                        break;
                    case '0' or '#' or '?':
                        tokens.Add(new Token(Code.Digit, c.ToString()));
                        break;
                    case '.':
                        tokens.Add(new Token(Code.Point, "."));
                        break;
                    case ',':
                        tokens.Add(new Token(Code.Comma, ","));
                        break;
                    case '%':
                        tokens.Add(new Token(Code.Percent, "%"));
                        break;
                    case '/':
                        tokens.Add(new Token(Code.Slash, "/"));
                        break;
                    case '@':
                        tokens.Add(new Token(Code.At));
                        break;
                    case 'E' or 'e' when i + 1 < text.Length && text[i + 1] is '+' or '-':
                        tokens.Add(new Token(Code.Exponent, text.Substring(i, 2)));
                        i++;
                        break;
                    case 'G' or 'g' when string.Compare(text, i, "General", 0, 7, StringComparison.OrdinalIgnoreCase) == 0:
                        tokens.Add(new Token(Code.General));
                        i += 6;
                        break;
                    case 'A' or 'a' when string.Compare(text, i, "AM/PM", 0, 5, StringComparison.OrdinalIgnoreCase) == 0:
                        tokens.Add(new Token(Code.AmPm, text.Substring(i, 5)));
                        i += 4;
                        break;
                    case 'A' or 'a' when string.Compare(text, i, "A/P", 0, 3, StringComparison.OrdinalIgnoreCase) == 0:
                        tokens.Add(new Token(Code.AmPm, text.Substring(i, 3)));
                        i += 2;
                        break;
                    case 'y' or 'Y' or 'm' or 'M' or 'd' or 'D' or 'h' or 'H' or 's' or 'S':
                        var lower = char.ToLowerInvariant(c);
                        var run = 1;
                        while (i + run < text.Length && char.ToLowerInvariant(text[i + run]) == lower)
                            run++;
                        var dateCode = lower switch
                        {
                            'y' => Code.Year,
                            'm' => Code.Month,
                            'd' => Code.Day,
                            'h' => Code.Hour,
                            _ => Code.Second,
                        };
                        tokens.Add(new Token(dateCode, Length: run));
                        i += run - 1;
                        break;
                    default:
                        tokens.Add(new Token(Code.Literal, c.ToString()));
                        break;
                }
            }

            return ResolveDateCodes(tokens) ? new Section(tokens, condition) : null;
        }

        // In a date section: m next to h or s is minutes, ".0" after seconds is fractions of a
        // second, and the number codes are plain characters. A digit placeholder is an error.
        private static bool ResolveDateCodes(List<Token> tokens)
        {
            var isDate = tokens.Exists(t => t.Code >= Code.Year);
            if (!isDate)
                return true;

            for (var i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];
                switch (token.Code)
                {
                    case Code.Month when token.Length <= 2 && (Previous(tokens, i) is Code.Hour or Code.ElapsedHours || Next(tokens, i) is Code.Second):
                        tokens[i] = new Token(Code.Minute, Length: token.Length);
                        break;
                    case Code.Point when i > 0 && tokens[i - 1].Code is Code.Second or Code.ElapsedSeconds
                                         && i + 1 < tokens.Count && tokens[i + 1] is { Code: Code.Digit, Text: "0" }:
                        var count = 0;
                        while (i + 1 + count < tokens.Count && tokens[i + 1 + count] is { Code: Code.Digit, Text: "0" })
                            count++;
                        tokens.RemoveRange(i + 1, count);
                        tokens[i] = new Token(Code.SecondFraction, Length: Math.Min(count, 3));
                        break;
                    case Code.Point or Code.Comma or Code.Percent or Code.Slash:
                        tokens[i] = new Token(Code.Literal, token.Text);
                        break;
                    case Code.Digit or Code.Exponent or Code.General:
                        return false;
                }
            }

            return true;
        }

        private static Code? Previous(List<Token> tokens, int index)
        {
            for (var i = index - 1; i >= 0; i--)
            {
                if (tokens[i].Code >= Code.Year)
                    return tokens[i].Code;
            }

            return null;
        }

        private static Code? Next(List<Token> tokens, int index)
        {
            for (var i = index + 1; i < tokens.Count; i++)
            {
                if (tokens[i].Code >= Code.Year)
                    return tokens[i].Code;
            }

            return null;
        }

        private static bool IsRun(string content, out char letter)
        {
            letter = char.ToLowerInvariant(content[0]);
            foreach (var c in content)
            {
                if (char.ToLowerInvariant(c) != letter)
                    return false;
            }

            return true;
        }

        private static bool IsLiteralDigit(Token token) => token.Code == Code.Literal && token.Text.Length == 1 && char.IsAsciiDigit(token.Text[0]);

        public string FormatText(string text)
        {
            var sb = new StringBuilder();
            foreach (var token in _tokens)
            {
                if (token.Code == Code.At)
                    sb.Append(text);
                else if (token.Code is Code.Literal or Code.Point or Code.Comma or Code.Percent or Code.Slash)
                    sb.Append(token.Text);
            }

            return sb.ToString();
        }

        // The value arrives signed when the section shows the sign itself (a lone section, or one
        // picked by a condition) and as a magnitude for the negative section.
        public string? Format(double value, CultureInfo culture, DateSystem dateSystem)
        {
            if (_isDate)
                return FormatDate(value, culture, dateSystem);

            var body = new StringBuilder();
            var abs = Math.Abs(value);
            bool nonZero;
            if (_exponent >= 0)
                nonZero = FormatScientific(body, abs, culture.NumberFormat);
            else if (_slash >= 0)
                nonZero = FormatFraction(body, abs, culture.NumberFormat);
            else
                nonZero = FormatPlain(body, abs, culture.NumberFormat);

            if (value < 0 && nonZero)
                body.Insert(0, '-');
            return body.ToString();
        }

        // Literals, percent signs and the General and @ codes, which appear among the digits.
        private bool AppendOther(StringBuilder sb, Token token, double abs, NumberFormatInfo nfi)
        {
            switch (token.Code)
            {
                case Code.Literal or Code.Percent or Code.Slash:
                    sb.Append(token.Text);
                    return true;
                case Code.General:
                    AppendGeneral(sb, abs * Math.Pow(100, _percent), nfi.NumberDecimalSeparator);
                    return true;
                case Code.At:
                    sb.Append(NumberText.FormatGeneral(abs, CultureInfo.InvariantCulture).Replace(".", nfi.NumberDecimalSeparator, StringComparison.Ordinal));
                    return true;
                case Code.Point:
                    sb.Append(nfi.NumberDecimalSeparator);
                    return true;
                case Code.Comma:
                    return true;
                default:
                    return false;
            }
        }

        private bool FormatPlain(StringBuilder sb, double abs, NumberFormatInfo nfi)
        {
            var number = DecimalDigits.Of(abs).Shift(2 * _percent - 3 * _scale).Round(_decimals);
            var integerEnd = _point >= 0 ? _point : _tokens.Count;
            AppendInteger(sb, 0, integerEnd, number.IntegerPart(), abs, nfi, beforePoint: _point >= 0 && _hasDigits);
            if (_point >= 0)
            {
                AppendOther(sb, _tokens[_point], abs, nfi);
                AppendFraction(sb, _point + 1, _tokens.Count, number.FractionPart(_decimals), abs, nfi);
            }

            return _hasDigits ? !number.IsZero : abs != 0;
        }

        private bool FormatScientific(StringBuilder sb, double abs, NumberFormatInfo nfi)
        {
            var integerEnd = _point >= 0 ? _point : _exponent;
            int places = 0, hashes = 0;
            for (var i = 0; i < integerEnd; i++)
            {
                if (_tokens[i].Code != Code.Digit)
                    continue;
                places++;
                if (_tokens[i].Placeholder == '#')
                    hashes++;
            }

            // "##0.0E+0" keeps the exponent a multiple of the integer places (engineering
            // notation); otherwise the integer places are all filled.
            var engineering = hashes > 0 && places > 1;
            var leading = Math.Max(places, 1);
            var digits = DecimalDigits.Of(abs).Shift(2 * _percent);
            var exponent = 0;
            var mantissa = digits;
            if (!digits.IsZero)
            {
                var magnitude = digits.Point - 1;
                exponent = engineering ? (int)Math.Floor(magnitude / (double)places) * places : magnitude - leading + 1;
                mantissa = digits.Shift(-exponent).Round(_decimals);
                if (mantissa.Point > (engineering ? places : leading))
                {
                    exponent += engineering ? places : 1;
                    mantissa = digits.Shift(-exponent).Round(_decimals);
                }
            }

            AppendInteger(sb, 0, integerEnd, mantissa.IntegerPart(), abs, nfi, beforePoint: true);
            if (_point >= 0)
            {
                AppendOther(sb, _tokens[_point], abs, nfi);
                AppendFraction(sb, _point + 1, _exponent, mantissa.FractionPart(_decimals), abs, nfi);
            }

            var marker = _tokens[_exponent].Text;
            sb.Append(marker[0]);
            if (exponent < 0)
                sb.Append('-');
            else if (marker[1] == '+')
                sb.Append('+');

            var exponentPlaces = 0;
            for (var i = _exponent + 1; i < _tokens.Count; i++)
            {
                if (_tokens[i].Code == Code.Digit && _tokens[i].Placeholder == '0')
                    exponentPlaces++;
            }

            var exponentDigits = Math.Abs(exponent).ToString(CultureInfo.InvariantCulture).PadLeft(exponentPlaces, '0');
            var written = false;
            for (var i = _exponent + 1; i < _tokens.Count; i++)
            {
                if (_tokens[i].Code == Code.Digit)
                {
                    if (!written)
                        sb.Append(exponentDigits);
                    written = true;
                }
                else
                {
                    AppendOther(sb, _tokens[i], abs, nfi);
                }
            }

            if (!written)
                sb.Append(exponentDigits);
            return !mantissa.IsZero;
        }

        // "# ?/?", "?/8", "# ??/100": an optional whole part, a numerator, a denominator that is
        // either placeholders (the best fraction with that many digits) or a fixed number.
        private bool FormatFraction(StringBuilder sb, double abs, NumberFormatInfo nfi)
        {
            var numeratorStart = _slash;
            while (numeratorStart > 0 && _tokens[numeratorStart - 1].Code == Code.Digit)
                numeratorStart--;
            var wholeEnd = numeratorStart;
            while (wholeEnd > 0 && _tokens[wholeEnd - 1].Code != Code.Digit)
                wholeEnd--;
            var hasWhole = wholeEnd > 0;

            var denominatorEnd = _slash + 1;
            long fixedDenominator = 0;
            var denominatorPlaces = 0;
            if (IsLiteralDigit(_tokens[denominatorEnd]))
            {
                // A fixed denominator such as 100 reads as the literal 1 and two '0' placeholders.
                while (denominatorEnd < _tokens.Count && fixedDenominator < 1_000_000_000
                       && (IsLiteralDigit(_tokens[denominatorEnd]) || _tokens[denominatorEnd] is { Code: Code.Digit, Text: "0" }))
                    fixedDenominator = fixedDenominator * 10 + (_tokens[denominatorEnd++].Text[0] - '0');
            }
            else
            {
                while (denominatorEnd < _tokens.Count && _tokens[denominatorEnd].Code == Code.Digit)
                {
                    denominatorEnd++;
                    denominatorPlaces++;
                }
            }

            var scaled = abs * Math.Pow(100, _percent);
            var whole = hasWhole ? Math.Floor(scaled) : 0;
            var rest = scaled - whole;
            long numerator, denominator;
            if (fixedDenominator > 0)
            {
                denominator = fixedDenominator;
                numerator = (long)Math.Round(rest * denominator, MidpointRounding.AwayFromZero);
            }
            else
            {
                (numerator, denominator) = Approximate(rest, (long)Math.Pow(10, Math.Min(denominatorPlaces, 9)) - 1);
            }

            if (hasWhole && numerator == denominator)
            {
                whole++;
                numerator = 0;
            }

            var showFraction = !(hasWhole && numerator == 0);
            var wholeDigits = whole == 0 ? (showFraction ? "" : "0") : DecimalDigits.Of(whole).IntegerPart();
            if (hasWhole)
                AppendInteger(sb, 0, wholeEnd, wholeDigits, abs, nfi, beforePoint: false);
            for (var i = wholeEnd; i < numeratorStart; i++)
                AppendOther(sb, _tokens[i], abs, nfi);

            if (showFraction)
            {
                AppendInteger(sb, numeratorStart, _slash, numerator.ToString(CultureInfo.InvariantCulture), abs, nfi, beforePoint: false, grouping: false);
                sb.Append('/');
                var denominatorText = denominator.ToString(CultureInfo.InvariantCulture);
                if (fixedDenominator > 0)
                {
                    sb.Append(denominatorText);
                }
                else
                {
                    // The denominator is left-aligned: '?' places after its digits become spaces.
                    var padding = 0;
                    for (var i = _slash + 1 + denominatorText.Length; i < denominatorEnd; i++)
                        padding += _tokens[i].Placeholder == '?' ? 1 : 0;
                    sb.Append(denominatorText).Append(' ', padding);
                }
            }
            else
            {
                // A whole number keeps the width of the fraction it does not show.
                for (var i = numeratorStart; i < denominatorEnd; i++)
                {
                    if (_tokens[i].Code != Code.Digit || _tokens[i].Placeholder == '?')
                        sb.Append(' ');
                }
            }

            for (var i = denominatorEnd; i < _tokens.Count; i++)
                AppendOther(sb, _tokens[i], abs, nfi);
            return whole != 0 || numerator != 0;
        }

        // The closest fraction with a denominator up to max, by continued fractions.
        private static (long Numerator, long Denominator) Approximate(double x, long max)
        {
            if (max < 1)
                max = 1;
            long p0 = 0, q0 = 1, p1 = 1, q1 = 0;
            var rest = x;
            for (var step = 0; step < 64; step++)
            {
                var a = (long)Math.Floor(rest);
                var q2 = q0 + a * q1;
                if (q2 > max)
                    break;
                var p2 = p0 + a * p1;
                (p0, q0, p1, q1) = (p1, q1, p2, q2);
                var fraction = rest - a;
                if (fraction < 1e-12)
                    break;
                rest = 1 / fraction;
            }

            // Also consider the best semiconvergent below the limit.
            if (q1 == 0)
                return (0, 1);
            var k = (max - q0) / q1;
            long ps = p0 + k * p1, qs = q0 + k * q1;
            return Math.Abs(x - (double)ps / qs) < Math.Abs(x - (double)p1 / q1) ? (ps, qs) : (p1, q1);
        }

        // Writes integer digits right-aligned into the placeholders of tokens [start, end):
        // missing digits become '0', ' ' or nothing by placeholder, extra digits go to the first one.
        private void AppendInteger(StringBuilder sb, int start, int end, string digits, double abs, NumberFormatInfo nfi, bool beforePoint, bool grouping = true)
        {
            var places = 0;
            for (var i = start; i < end; i++)
            {
                if (_tokens[i].Code == Code.Digit)
                    places++;
            }

            var cells = new string[places];
            var place = places;
            var used = 0;
            for (var i = end - 1; i >= start; i--)
            {
                if (_tokens[i].Code != Code.Digit)
                    continue;
                place--;
                if (used < digits.Length)
                {
                    cells[place] = digits[digits.Length - 1 - used].ToString();
                    used++;
                }
                else
                {
                    cells[place] = _tokens[i].Placeholder switch { '0' => "0", '?' => " ", _ => "" };
                }
            }

            var extra = digits[..(digits.Length - used)];
            var separator = grouping && _grouping ? nfi.NumberGroupSeparator : "";
            var remaining = extra.Length;
            foreach (var cell in cells)
            {
                if (cell.Length == 1 && char.IsAsciiDigit(cell[0]))
                    remaining++;
            }

            void Append(char c)
            {
                sb.Append(c);
                if (!char.IsAsciiDigit(c))
                    return;
                remaining--;
                if (separator.Length > 0 && remaining > 0 && remaining % 3 == 0)
                    sb.Append(separator);
            }

            // A format without integer places (".00") still shows the integer digits.
            if (places == 0 && beforePoint)
            {
                foreach (var c in digits)
                    Append(c);
            }

            place = 0;
            for (var i = start; i < end; i++)
            {
                var token = _tokens[i];
                if (token.Code != Code.Digit)
                {
                    AppendOther(sb, token, abs, nfi);
                    continue;
                }

                if (place == 0)
                {
                    foreach (var c in extra)
                        Append(c);
                }

                foreach (var c in cells[place++])
                    Append(c);
            }
        }

        // Writes fraction digits left to right; trailing zeros under '#' vanish and under '?' become spaces.
        private void AppendFraction(StringBuilder sb, int start, int end, string digits, double abs, NumberFormatInfo nfi)
        {
            var placeholders = new List<char>();
            for (var i = start; i < end; i++)
            {
                if (_tokens[i].Code == Code.Digit)
                    placeholders.Add(_tokens[i].Placeholder);
            }

            var significant = placeholders.Count;
            while (significant > 0 && digits[significant - 1] == '0' && placeholders[significant - 1] != '0')
                significant--;

            var place = 0;
            for (var i = start; i < end; i++)
            {
                var token = _tokens[i];
                if (token.Code != Code.Digit)
                {
                    AppendOther(sb, token, abs, nfi);
                    continue;
                }

                if (place < significant)
                    sb.Append(digits[place]);
                else if (token.Placeholder == '?')
                    sb.Append(' ');
                place++;
            }
        }

        private string? FormatDate(double serial, CultureInfo culture, DateSystem dateSystem)
        {
            if (!(serial >= 0) || serial >= DateSerial.MaxSerial(dateSystem) + 1)
                return null;

            // Round to the precision shown, carrying into the date: 23:59:59.7 is the next day's 00:00:00.
            var unitsPerSecond = Math.Pow(10, _secondDigits);
            var units = Math.Round(serial * 86400 * unitsPerSecond, MidpointRounding.AwayFromZero);
            var unitsPerDay = 86400 * unitsPerSecond;
            var day = Math.Floor(units / unitsPerDay);
            var timeUnits = units - day * unitsPerDay;
            var totalSeconds = Math.Floor(units / unitsPerSecond);
            var secondOfDay = (long)Math.Floor(timeUnits / unitsPerSecond);
            var fraction = (long)(timeUnits - secondOfDay * unitsPerSecond);

            if (!DateSerial.TryToDate(day, dateSystem, out var year, out var month, out var dayOfMonth))
            {
                // Rounding up from 9999-12-31 23:59:59.5 leaves the calendar.
                if (!DateSerial.TryToDate(Math.Floor(serial), dateSystem, out year, out month, out dayOfMonth))
                    return null;
            }

            var hour = (int)(secondOfDay / 3600);
            var minute = (int)(secondOfDay / 60 % 60);
            var second = (int)(secondOfDay % 60);
            var names = culture.DateTimeFormat;
            var weekday = (int)((day + (dateSystem == DateSystem.Date1904 ? 5 : 6)) % 7);

            var sb = new StringBuilder();
            foreach (var token in _tokens)
            {
                switch (token.Code)
                {
                    case Code.Year:
                        sb.Append(token.Length <= 2
                            ? (year % 100).ToString("00", CultureInfo.InvariantCulture)
                            : year.ToString("0000", CultureInfo.InvariantCulture));
                        break;
                    case Code.Month:
                        sb.Append(token.Length switch
                        {
                            1 => month.ToString(CultureInfo.InvariantCulture),
                            2 => month.ToString("00", CultureInfo.InvariantCulture),
                            3 => names.AbbreviatedMonthNames[month - 1],
                            5 => names.MonthNames[month - 1][..1],
                            _ => names.MonthNames[month - 1],
                        });
                        break;
                    case Code.Day:
                        sb.Append(token.Length switch
                        {
                            1 => dayOfMonth.ToString(CultureInfo.InvariantCulture),
                            2 => dayOfMonth.ToString("00", CultureInfo.InvariantCulture),
                            3 => names.AbbreviatedDayNames[weekday],
                            _ => names.DayNames[weekday],
                        });
                        break;
                    case Code.Hour:
                        var shownHour = _twelveHour ? (hour % 12 == 0 ? 12 : hour % 12) : hour;
                        sb.Append(shownHour.ToString(token.Length >= 2 ? "00" : "0", CultureInfo.InvariantCulture));
                        break;
                    case Code.Minute:
                        sb.Append(minute.ToString(token.Length >= 2 ? "00" : "0", CultureInfo.InvariantCulture));
                        break;
                    case Code.Second:
                        sb.Append(second.ToString(token.Length >= 2 ? "00" : "0", CultureInfo.InvariantCulture));
                        break;
                    case Code.ElapsedHours:
                        sb.Append(Math.Floor(totalSeconds / 3600).ToString(new string('0', token.Length), CultureInfo.InvariantCulture));
                        break;
                    case Code.ElapsedMinutes:
                        sb.Append(Math.Floor(totalSeconds / 60).ToString(new string('0', token.Length), CultureInfo.InvariantCulture));
                        break;
                    case Code.ElapsedSeconds:
                        sb.Append(totalSeconds.ToString(new string('0', token.Length), CultureInfo.InvariantCulture));
                        break;
                    case Code.SecondFraction:
                        sb.Append(culture.NumberFormat.NumberDecimalSeparator)
                            .Append(fraction.ToString(CultureInfo.InvariantCulture).PadLeft(_secondDigits, '0'));
                        break;
                    case Code.AmPm:
                        var pm = hour >= 12;
                        sb.Append(token.Text.Length == 5
                            ? (pm ? token.Text[3..5] : token.Text[..2])
                            : (pm ? token.Text[2] : token.Text[0]));
                        break;
                    case Code.Literal:
                        sb.Append(token.Text);
                        break;
                }
            }

            return sb.ToString();
        }
    }

    // General in 11 characters: fixed notation from 1E-9 up to 11 integer digits, rounded to
    // what fits (at most 10 significant digits with a decimal point); scientific otherwise.
    private static void AppendGeneral(StringBuilder sb, double abs, string separator)
    {
        if (abs == 0)
        {
            sb.Append('0');
            return;
        }

        var digits = DecimalDigits.Of(abs);
        var magnitude = digits.Point - 1;
        if (magnitude is >= -9 and <= 10)
        {
            var decimals = magnitude < 0 ? 9 : 9 - magnitude;
            DecimalDigits rounded;
            if (magnitude <= -5)
            {
                // Tiny numbers stay fixed only if nothing is lost in 11 characters.
                rounded = digits.Round(12);
                if (rounded.Digits.Length - rounded.Point > 9)
                    rounded = default;
            }
            else
            {
                rounded = digits.Round(Math.Max(decimals, 0));
            }

            if (rounded.Digits is not null && rounded.Point <= 11)
            {
                var integer = rounded.IntegerPart();
                var fraction = rounded.FractionPart(Math.Max(rounded.Digits.Length - rounded.Point, 0));
                sb.Append(integer.Length == 0 ? "0" : integer);
                if (fraction.Length > 0)
                    sb.Append(separator).Append(fraction);
                return;
            }
        }

        var exponent = magnitude;
        var mantissa = digits.Shift(-exponent).Round(5);
        if (mantissa.Point > 1)
        {
            exponent++;
            mantissa = digits.Shift(-exponent).Round(5);
        }

        var mantissaFraction = mantissa.FractionPart(Math.Max(mantissa.Digits.Length - mantissa.Point, 0));
        sb.Append(mantissa.IntegerPart());
        if (mantissaFraction.Length > 0)
            sb.Append(separator).Append(mantissaFraction);
        sb.Append('E').Append(exponent < 0 ? '-' : '+')
            .Append(Math.Abs(exponent).ToString("00", CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// A non-negative number as at most 15 significant decimal digits with the position of the
    /// decimal point: 12.5 is ("125", 2), 0.004 is ("4", -2), zero is ("", 0).
    /// </summary>
    private readonly record struct DecimalDigits(string Digits, int Point)
    {
        public bool IsZero => Digits.Length == 0;

        public static DecimalDigits Of(double abs)
        {
            if (abs == 0)
                return new DecimalDigits("", 0);

            var text = abs.ToString("E" + (SignificantDigits - 1), CultureInfo.InvariantCulture);
            var e = text.IndexOf('E');
            var exponent = int.Parse(text.AsSpan(e + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
            var digits = string.Concat(text.AsSpan(0, 1), text.AsSpan(2, e - 2)).TrimEnd('0');
            return new DecimalDigits(digits, exponent + 1);
        }

        /// <summary>Multiplies by 10^places.</summary>
        public DecimalDigits Shift(int places) => IsZero ? this : this with { Point = Point + places };

        /// <summary>Rounds half away from zero to the given number of decimals.</summary>
        public DecimalDigits Round(int decimals)
        {
            var keep = Point + decimals;
            if (keep >= Digits.Length)
                return this;
            if (keep < 0)
                return new DecimalDigits("", 0);

            var up = Digits[keep] >= '5';
            var kept = Digits[..keep].ToCharArray();
            var point = Point;
            if (up)
            {
                var i = kept.Length - 1;
                while (i >= 0 && kept[i] == '9')
                    kept[i--] = '0';
                if (i >= 0)
                {
                    kept[i]++;
                }
                else
                {
                    point++;
                    return new DecimalDigits(("1" + new string(kept)).TrimEnd('0'), point);
                }
            }

            var result = new string(kept).TrimEnd('0');
            return result.Length == 0 ? new DecimalDigits("", 0) : new DecimalDigits(result, point);
        }

        public string IntegerPart()
        {
            if (Point <= 0 || IsZero)
                return "";
            return Point >= Digits.Length ? Digits + new string('0', Point - Digits.Length) : Digits[..Point];
        }

        /// <summary>Exactly <paramref name="places"/> digits after the point.</summary>
        public string FractionPart(int places)
        {
            if (places <= 0)
                return "";
            var sb = new StringBuilder(places);
            for (var i = 0; i < places; i++)
            {
                var index = Point + i;
                sb.Append(index >= 0 && index < Digits.Length ? Digits[index] : '0');
            }

            return sb.ToString();
        }
    }
}
