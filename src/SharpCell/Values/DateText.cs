using System;
using System.Collections.Generic;
using System.Globalization;

namespace SharpCell;

/// <summary>
/// Dates and times written as text, as Excel reads them when text meets arithmetic or a date
/// function: <c>"12:00"</c>, <c>"2020-01-01"</c>, <c>"1/2/2020 13:30"</c>. The result is a serial number.
/// </summary>
/// <remarks>
/// <para>Accepted, with spaces allowed around the separators:</para>
/// <list type="bullet">
/// <item>times: <c>h:mm</c>, <c>h:mm:ss</c>, <c>h:mm:ss.000</c>, <c>mm:ss.0</c>, each with an optional
/// AM/PM, and <c>h AM</c>. Hours may exceed 23 ("25:00" is 1.0417) unless AM/PM is given; minutes and
/// seconds stay below 60 ("23:59:60", exactly one day, is accepted);</item>
/// <item>numeric dates with <c>/</c>, <c>-</c> or the culture's date separator: year first when it has
/// four digits ("2020-01-31"), otherwise in the culture's order (month first for the invariant
/// culture, like English Excel: "1/31/2020"); two-digit years 00-29 are 2000-2029, 30-99 are 1930-1999;</item>
/// <item>dates with a month name, English or the culture's, full or abbreviated: "1-Jan-2020",
/// "1 January 2020", "Jan-2020" (the 1st); in month-first cultures also "Jan 1, 2020" and "January 1 2020";</item>
/// <item>a date followed by a time: "2020-01-31 13:30", "1/31/2020 1:30 PM".</item>
/// </list>
/// <para>
/// Dates must exist ("2/30/2020" is not one) between 1900 and 9999; 1900-02-29 exists in the 1900
/// system. Text without a year ("1-Jan", "3/15") means the current year, which needs the workbook
/// clock: plain coercion has none and does not read such text, DATEVALUE and TIMEVALUE pass the year.
/// </para>
/// </remarks>
internal static class DateText
{
    private static readonly string[] EnglishMonths =
    [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    ];

    // Excel reads hours up to 9999 in "9999:59".
    private const int MaxHours = 9999;

    private enum TokenKind
    {
        Number,
        Word,
        Separator,
        Space,
    }

    private enum DateOrder
    {
        MonthDayYear,
        DayMonthYear,
        YearMonthDay,
    }

    private readonly record struct Token(TokenKind Kind, int Start, int Length, char Separator);

    public static bool TryParse(string text, CultureInfo culture, DateSystem dateSystem, out double serial) =>
        TryParse(text, culture, dateSystem, currentYear: null, out serial);

    /// <param name="currentYear">The year of dates written without one; null rejects such text.</param>
    public static bool TryParse(string text, CultureInfo culture, DateSystem dateSystem, int? currentYear, out double serial)
    {
        serial = 0;
        var format = culture.DateTimeFormat;
        var dateSeparator = format.DateSeparator.Length == 1 ? format.DateSeparator[0] : '/';
        var timeSeparator = format.TimeSeparator.Length == 1 && format.TimeSeparator[0] != dateSeparator ? format.TimeSeparator[0] : ':';

        if (!TryTokenize(text, dateSeparator, timeSeparator, out var tokens) || tokens.Count == 0)
            return false;

        // The time part starts at the number before the first time separator, or before a final AM/PM.
        int? timeStart = null;
        for (var i = 0; i < tokens.Count; i++)
        {
            if (tokens[i].Kind == TokenKind.Separator && IsTimeSeparator(tokens[i].Separator, timeSeparator))
            {
                timeStart = i - 1;
                break;
            }
        }

        if (timeStart is null && tokens[^1].Kind == TokenKind.Word && TryMeridiem(text, tokens[^1], culture, out _))
        {
            var start = tokens.Count - 2;
            timeStart = start >= 0 && tokens[start].Kind == TokenKind.Space ? start - 1 : start;
        }

        var dateEnd = tokens.Count;
        double time = 0;
        if (timeStart is { } first)
        {
            if (first < 0 || tokens[first].Kind != TokenKind.Number
                || !TryParseTime(text, tokens, first, timeSeparator, culture, out time))
                return false;

            // A date before the time is separated from it by a space.
            dateEnd = first;
            if (dateEnd > 0)
            {
                if (tokens[dateEnd - 1].Kind != TokenKind.Space)
                    return false;
                dateEnd--;
            }
        }

        double date = 0;
        if (dateEnd > 0 && !TryParseDate(text, tokens, dateEnd, dateSeparator, culture, dateSystem, currentYear, out date))
            return false;

        serial = date + time;
        return serial < DateSerial.MaxSerial(dateSystem) + 1;
    }

    // Splits into numbers, words, separators and spaces. Spaces next to a separator are dropped:
    // "2026 - 01 - 01" reads as "2026-01-01" and "4 : 35" as "4:35".
    private static bool TryTokenize(string text, char dateSeparator, char timeSeparator, out List<Token> tokens)
    {
        tokens = [];
        var i = 0;
        while (i < text.Length)
        {
            var c = text[i];
            var start = i;
            if (char.IsAsciiDigit(c))
            {
                while (i < text.Length && char.IsAsciiDigit(text[i]))
                    i++;
                tokens.Add(new Token(TokenKind.Number, start, i - start, '\0'));
            }
            else if (char.IsLetter(c))
            {
                while (i < text.Length && char.IsLetter(text[i]))
                    i++;
                tokens.Add(new Token(TokenKind.Word, start, i - start, '\0'));
            }
            else if (char.IsWhiteSpace(c))
            {
                while (i < text.Length && char.IsWhiteSpace(text[i]))
                    i++;
                if (tokens.Count > 0 && tokens[^1].Kind != TokenKind.Separator)
                    tokens.Add(new Token(TokenKind.Space, start, i - start, ' '));
            }
            else if (c is '/' or '-' or ':' or ',' or '.' || c == dateSeparator || c == timeSeparator)
            {
                if (tokens.Count > 0 && tokens[^1].Kind == TokenKind.Space)
                    tokens.RemoveAt(tokens.Count - 1);
                tokens.Add(new Token(TokenKind.Separator, start, 1, c));
                i++;
            }
            else
            {
                return false;
            }
        }

        if (tokens.Count > 0 && tokens[^1].Kind == TokenKind.Space)
            tokens.RemoveAt(tokens.Count - 1);
        return true;
    }

    // h:mm, h:mm:ss, h:mm:ss.000, mm:ss.0 and h, each with an optional AM/PM (required for a bare hour).
    private static bool TryParseTime(string text, List<Token> tokens, int start, char timeSeparator, CultureInfo culture, out double fraction)
    {
        fraction = 0;
        Span<int> parts = stackalloc int[3];
        Span<int> digits = stackalloc int[3];
        var count = 0;
        var i = start;
        while (true)
        {
            if (count == 3 || !TryNumber(text, tokens[i], out parts[count], out digits[count]))
                return false;
            count++;
            i++;
            if (i + 1 < tokens.Count && tokens[i].Kind == TokenKind.Separator && IsTimeSeparator(tokens[i].Separator, timeSeparator)
                && tokens[i + 1].Kind == TokenKind.Number)
            {
                i++;
                continue;
            }

            break;
        }

        // Fractional seconds: "4:35:00.5", or "12:30.5" (minutes and seconds).
        double secondsFraction = 0;
        var hasFraction = false;
        if (count >= 2 && i + 1 < tokens.Count && tokens[i].Kind == TokenKind.Separator
            && (tokens[i].Separator == '.' || culture.NumberFormat.NumberDecimalSeparator == tokens[i].Separator.ToString())
            && tokens[i + 1].Kind == TokenKind.Number)
        {
            var fractionDigits = text.AsSpan(tokens[i + 1].Start, tokens[i + 1].Length);
            secondsFraction = double.Parse("0." + fractionDigits.ToString(), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture);
            hasFraction = true;
            i += 2;
        }

        bool? pm = null;
        if (i < tokens.Count && tokens[i].Kind == TokenKind.Space)
            i++;
        if (i < tokens.Count)
        {
            if (tokens[i].Kind != TokenKind.Word || !TryMeridiem(text, tokens[i], culture, out var isPm))
                return false;
            pm = isPm;
            i++;
        }

        if (i != tokens.Count)
            return false;

        int hours, minutes, seconds;
        if (count == 1)
        {
            if (pm is null)
                return false;
            (hours, minutes, seconds) = (parts[0], 0, 0);
        }
        else if (count == 2 && hasFraction)
        {
            if (pm is not null || digits[1] > 2 || parts[1] > 59 || parts[0] > MaxHours * 60)
                return false;
            (hours, minutes, seconds) = (0, parts[0], parts[1]);
        }
        else
        {
            if (digits[1] > 2 || (count == 3 && digits[2] > 2) || parts[1] > 59 || (count == 3 && parts[2] > 60))
                return false;
            (hours, minutes, seconds) = (parts[0], parts[1], count == 3 ? parts[2] : 0);
        }

        if (pm is { } afternoon)
        {
            if (hours > 12)
                return false;
            hours = hours % 12 + (afternoon ? 12 : 0);
        }
        else if (hours > MaxHours)
        {
            return false;
        }

        fraction = (hours * 3600.0 + minutes * 60.0 + seconds + secondsFraction) / 86400;
        return true;
    }

    private static bool TryParseDate(string text, List<Token> tokens, int end, char dateSeparator, CultureInfo culture,
        DateSystem system, int? currentYear, out double serial)
    {
        serial = 0;

        // Items alternate with separators: item (separator item)*, at most three items.
        Span<int> items = stackalloc int[3];
        Span<char> separators = stackalloc char[2];
        var count = 0;
        for (var i = 0; i < end; i++)
        {
            var token = tokens[i];
            if (i % 2 == 0)
            {
                if (token.Kind is not (TokenKind.Number or TokenKind.Word) || count == 3)
                    return false;
                items[count++] = i;
            }
            else
            {
                if (token.Kind is not (TokenKind.Separator or TokenKind.Space))
                    return false;
                separators[count - 1] = token.Separator;
            }
        }

        if (count < 2 || end % 2 == 0)
            return false;

        var order = OrderOf(culture);
        var wordAt = -1;
        for (var i = 0; i < count; i++)
        {
            if (tokens[items[i]].Kind == TokenKind.Word)
            {
                if (wordAt >= 0)
                    return false;
                wordAt = i;
            }
        }

        int year, month, day;
        if (wordAt < 0)
        {
            // Numbers only: one separator throughout, '/', '-' or the culture's.
            var separator = separators[0];
            if (separator is not ('/' or '-') && separator != dateSeparator)
                return false;
            if (count == 3 && separators[1] != separator)
                return false;
            if (!TryNumericDate(text, tokens, items[..count], order, system, currentYear, out year, out month, out day))
                return false;
        }
        else
        {
            if (!TryMonth(text, tokens[items[wordAt]], culture, out month))
                return false;
            for (var i = 0; i < count - 1; i++)
            {
                // A comma only between the day and the year of "Jan 1, 2020".
                var separator = separators[i];
                var comma = separator == ',' && wordAt == 0 && count == 3 && i == 1;
                if (!comma && separator is not ('/' or '-' or ' ') && separator != dateSeparator)
                    return false;
            }

            if (!TryNamedMonthDate(text, tokens, items[..count], wordAt, month, order, system, currentYear, out year, out day))
                return false;
        }

        return DateSerial.TryFromCalendar(year, month, day, system, out serial);
    }

    private static bool TryNumericDate(string text, List<Token> tokens, ReadOnlySpan<int> items, DateOrder order,
        DateSystem system, int? currentYear, out int year, out int month, out int day)
    {
        year = month = day = 0;
        Span<int> values = stackalloc int[3];
        Span<int> digits = stackalloc int[3];
        for (var i = 0; i < items.Length; i++)
        {
            if (!TryNumber(text, tokens[items[i]], out values[i], out digits[i]))
                return false;
        }

        if (items.Length == 3)
        {
            // A four-digit year first is read year-month-day in every culture.
            if (digits[0] == 4)
                order = DateOrder.YearMonthDay;
            else if (order == DateOrder.YearMonthDay && digits[2] == 4)
                order = DateOrder.MonthDayYear;

            var (y, m, d) = order switch
            {
                DateOrder.MonthDayYear => (2, 0, 1),
                DateOrder.DayMonthYear => (2, 1, 0),
                _ => (0, 1, 2),
            };
            if (digits[m] > 2 || digits[d] > 2 || !TryYear(values[y], digits[y], out year))
                return false;
            (month, day) = (values[m], values[d]);
            return true;
        }

        // Two numbers: a four-digit year first is year-month; otherwise month and day of the current
        // year, or month and year when the second number cannot be a day ("1/2020", "1/45").
        if (digits[0] == 4)
        {
            if (!TryYear(values[0], digits[0], out year) || digits[1] > 2)
                return false;
            (month, day) = (values[1], 1);
            return true;
        }

        if (order == DateOrder.DayMonthYear)
        {
            if (digits[1] <= 2 && IsDayOfMonth(values[0], digits[0], values[1], system))
                return WithCurrentYear(currentYear, values[1], values[0], out year, out month, out day);
        }
        else if (digits[0] <= 2 && IsDayOfMonth(values[1], digits[1], values[0], system))
        {
            return WithCurrentYear(currentYear, values[0], values[1], out year, out month, out day);
        }

        if (digits[0] > 2 || !TryYear(values[1], digits[1], out year))
            return false;
        (month, day) = (values[0], 1);
        return true;
    }

    // "1-Jan-2020", "1 Jan", "2020-Jan-01", "Jan-2020", and in month-first cultures "Jan 1, 2020" and "Jan 1".
    private static bool TryNamedMonthDate(string text, List<Token> tokens, ReadOnlySpan<int> items, int wordAt, int month,
        DateOrder order, DateSystem system, int? currentYear, out int year, out int day)
    {
        year = day = 0;
        Span<int> values = stackalloc int[3];
        Span<int> digits = stackalloc int[3];
        for (var i = 0; i < items.Length; i++)
        {
            if (i != wordAt && !TryNumber(text, tokens[items[i]], out values[i], out digits[i]))
                return false;
        }

        if (wordAt == 1)
        {
            if (items.Length == 2)
                return digits[0] <= 2 && WithCurrentYear(currentYear, month, values[0], out year, out _, out day);
            if (digits[0] == 4)
            {
                day = values[2];
                return digits[2] <= 2 && TryYear(values[0], digits[0], out year);
            }

            day = values[0];
            return digits[0] <= 2 && TryYear(values[2], digits[2], out year);
        }

        if (wordAt != 0)
            return false;

        if (items.Length == 2)
        {
            if (order == DateOrder.MonthDayYear && digits[1] <= 2 && IsDayOfMonth(values[1], digits[1], month, system))
                return WithCurrentYear(currentYear, month, values[1], out year, out _, out day);
            day = 1;
            return TryYear(values[1], digits[1], out year);
        }

        day = values[1];
        return order == DateOrder.MonthDayYear && digits[1] <= 2 && TryYear(values[2], digits[2], out year);
    }

    private static bool WithCurrentYear(int? currentYear, int month, int day, out int year, out int monthOut, out int dayOut)
    {
        (year, monthOut, dayOut) = (currentYear ?? 0, month, day);
        return currentYear is not null;
    }

    // A day that exists in the month in some year (February 29 included).
    private static bool IsDayOfMonth(int day, int digits, int month, DateSystem system) =>
        digits <= 2 && month is >= 1 and <= 12 && day >= 1 && day <= DateSerial.DaysInMonth(2000, month, system);

    // Four digits as written (1900-9999); one or two digits: 00-29 are 2000-2029, 30-99 are 1930-1999.
    private static bool TryYear(int value, int digits, out int year)
    {
        year = 0;
        if (digits == 4)
            year = value;
        else if (digits <= 2)
            year = value < 30 ? 2000 + value : 1900 + value;
        else
            return false;
        return year >= 1900;
    }

    private static bool TryNumber(string text, Token token, out int value, out int digits)
    {
        value = 0;
        digits = token.Length;
        if (token.Kind != TokenKind.Number)
            return false;

        var span = text.AsSpan(token.Start, token.Length).TrimStart('0');
        if (span.Length > 9)
            return false;
        value = span.Length == 0 ? 0 : int.Parse(span, NumberStyles.None, CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryMonth(string text, Token token, CultureInfo culture, out int month)
    {
        var word = text.AsSpan(token.Start, token.Length);
        for (var i = 0; i < 12; i++)
        {
            var english = EnglishMonths[i];
            if (word.Equals(english, StringComparison.OrdinalIgnoreCase)
                || word.Equals(english.AsSpan(0, 3), StringComparison.OrdinalIgnoreCase))
            {
                month = i + 1;
                return true;
            }
        }

        if (word.Equals("Sept", StringComparison.OrdinalIgnoreCase))
        {
            month = 9;
            return true;
        }

        var format = culture.DateTimeFormat;
        return TryName(word, format.MonthNames, culture, out month)
            || TryName(word, format.AbbreviatedMonthNames, culture, out month)
            || TryName(word, format.MonthGenitiveNames, culture, out month)
            || TryName(word, format.AbbreviatedMonthGenitiveNames, culture, out month);
    }

    private static bool TryName(ReadOnlySpan<char> word, string[] names, CultureInfo culture, out int month)
    {
        for (var i = 0; i < 12 && i < names.Length; i++)
        {
            var name = names[i].AsSpan().TrimEnd('.');
            if (name.Length > 0 && culture.CompareInfo.Compare(word, name, CompareOptions.IgnoreCase) == 0)
            {
                month = i + 1;
                return true;
            }
        }

        month = 0;
        return false;
    }

    private static bool TryMeridiem(string text, Token token, CultureInfo culture, out bool pm)
    {
        var word = text.AsSpan(token.Start, token.Length);
        var format = culture.DateTimeFormat;
        if (word.Equals("PM", StringComparison.OrdinalIgnoreCase)
            || (format.PMDesignator.Length > 0 && culture.CompareInfo.Compare(word, format.PMDesignator, CompareOptions.IgnoreCase) == 0))
        {
            pm = true;
            return true;
        }

        pm = false;
        return word.Equals("AM", StringComparison.OrdinalIgnoreCase)
            || (format.AMDesignator.Length > 0 && culture.CompareInfo.Compare(word, format.AMDesignator, CompareOptions.IgnoreCase) == 0);
    }

    private static bool IsTimeSeparator(char c, char timeSeparator) => c == ':' || c == timeSeparator;

    // The order of day, month and year in the culture's short date pattern.
    private static DateOrder OrderOf(CultureInfo culture)
    {
        var pattern = culture.DateTimeFormat.ShortDatePattern;
        var d = pattern.IndexOf('d');
        var m = pattern.IndexOf('M');
        var y = pattern.IndexOf('y');
        if (y >= 0 && (m < 0 || y < m) && (d < 0 || y < d))
            return DateOrder.YearMonthDay;
        return d >= 0 && m >= 0 && d < m ? DateOrder.DayMonthYear : DateOrder.MonthDayYear;
    }
}
