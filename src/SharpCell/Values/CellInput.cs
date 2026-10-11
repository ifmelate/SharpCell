using System;
using System.Globalization;
using SharpCell.Functions;

namespace SharpCell;

/// <summary>
/// Text typed into a cell, read as Excel reads it, and the text Excel shows to edit a cell. Numbers
/// and dates are read by the same rules as text meeting arithmetic and DATEVALUE, in the workbook's
/// culture; the shape of what was typed gives a cell in the General format its number format.
/// </summary>
internal static class CellInput
{
    /// <summary>What typing a text puts in a cell.</summary>
    /// <param name="Value">The constant, when there is no formula.</param>
    /// <param name="Formula">The formula, starting with <c>=</c>.</param>
    /// <param name="NumberFormat">The format a cell in the General format takes, or null to keep its format.</param>
    internal readonly record struct Entry(CellValue Value, string? Formula, string? NumberFormat);

    public static Entry Read(string text, CultureInfo culture, DateSystem dateSystem, int currentYear)
    {
        if (text.Length == 0)
            return new Entry(CellValue.Empty, null, null);
        if (text[0] == '\'')
            return new Entry(CellValue.Text(text[1..]), null, null);
        if (text[0] == '=' && text.Length > 1)
            return new Entry(CellValue.Empty, text, null);

        var trimmed = text.Trim();
        if (trimmed.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
            return new Entry(CellValue.Boolean(true), null, null);
        if (trimmed.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
            return new Entry(CellValue.Boolean(false), null, null);
        if (ErrorKinds.TryParse(trimmed, out var error))
            return new Entry(CellValue.Error(error), null, null);

        if (SharpCell.NumberText.TryParse(text, culture, out var number))
            return new Entry(CellValue.Number(number), null, NumberShape(trimmed, culture));
        if (DateText.TryParse(trimmed, culture, dateSystem, currentYear, out var serial))
            return new Entry(CellValue.Number(serial), null, DateShape(trimmed, culture));
        return new Entry(CellValue.Text(text), null, null);
    }

    /// <summary>The text Excel shows to edit a cell, which typed back gives the same content.</summary>
    public static string EditText(Cell cell)
    {
        if (cell.Formula is { } formula)
            return formula;
        var value = cell.Value;
        var workbook = cell.Worksheet.Workbook;
        var culture = workbook.Culture;
        switch (value.Kind)
        {
            case CellValueKind.Boolean:
                return value.AsBoolean() ? "TRUE" : "FALSE";
            case CellValueKind.Error:
                return value.AsError().ToText();
            case CellValueKind.Text:
                // Text that would read as something else is kept text by an apostrophe.
                var text = value.AsText();
                var read = Read(text, culture, workbook.DateSystem, 2000);
                return read.Formula is null && read.Value.Kind == CellValueKind.Text && read.Value.AsText() == text ? text : "'" + text;
            case CellValueKind.Number:
                return NumberEditText(value.AsNumber(), cell.Worksheet.FormatAt(cell.Row, cell.Column), culture, workbook.DateSystem);
            default:
                return "";
        }
    }

    private static string NumberEditText(double number, NumberFormat? format, CultureInfo culture, DateSystem dateSystem)
    {
        if (format is null)
            return SharpCell.NumberText.FormatGeneral(number, culture);
        if (format.ShowsPercent)
            return SharpCell.NumberText.FormatGeneral(number * 100, culture) + culture.NumberFormat.PercentSymbol;
        if (!format.ShowsDate && !format.ShowsTime)
            return SharpCell.NumberText.FormatGeneral(number, culture);

        // Dates and times as the culture writes them in full: the date when there is one, the time
        // when there is one, both for a date with a time of day.
        var day = Math.Floor(number);
        var showDate = format.ShowsDate || day >= 1;
        var showTime = format.ShowsTime || number != day;
        var code = showDate && showTime ? DateCode(culture) + " " + TimeCode(culture, seconds: true)
            : showDate ? DateCode(culture)
            : TimeCode(culture, seconds: true);
        return SharpCell.Functions.NumberFormat.TryParse(code, out var edit) && edit!.Format(number, culture, dateSystem) is { } shown
            ? shown
            : SharpCell.NumberText.FormatGeneral(number, culture);
    }

    // The format a number typed in a General cell takes from how it was written.
    private static string? NumberShape(string text, CultureInfo culture)
    {
        var format = culture.NumberFormat;
        var decimals = text.Contains(format.NumberDecimalSeparator, StringComparison.Ordinal);
        if (text.EndsWith('%'))
            return decimals ? "0.00%" : "0%";

        var currency = culture.Name.Length == 0 ? "$" : format.CurrencySymbol;
        var digits = decimals ? "#,##0.00" : "#,##0";
        if (currency.Length > 0 && text.Contains(currency, StringComparison.Ordinal))
        {
            var symbol = "\"" + currency + "\"";
            return (culture.Name.Length == 0 ? 0 : format.CurrencyPositivePattern) switch
            {
                1 => digits + symbol,
                2 => symbol + " " + digits,
                3 => digits + " " + symbol,
                _ => symbol + digits,
            };
        }

        if (format.NumberGroupSeparator.Length > 0 && text.Contains(format.NumberGroupSeparator, StringComparison.Ordinal))
            return digits;
        if (text.Contains('E', StringComparison.OrdinalIgnoreCase))
            return "0.00E+00";
        return null;
    }

    // The format a date or time typed in a General cell takes: the culture's short date, a time with
    // seconds when they were typed and AM/PM when it was.
    private static string DateShape(string text, CultureInfo culture)
    {
        var format = culture.DateTimeFormat;
        var timeSeparator = format.TimeSeparator.Length == 1 ? format.TimeSeparator[0] : ':';
        var separators = text.AsSpan().Count(timeSeparator) + (timeSeparator == ':' ? 0 : text.AsSpan().Count(':'));
        var twelveHour = HasMeridiem(text, format);
        var hasTime = separators > 0 || twelveHour;
        var hasDate = !hasTime || HasDatePart(text, format);
        var time = "h:mm" + (separators >= 2 ? ":ss" : "") + (twelveHour ? " AM/PM" : "");
        return hasDate && hasTime ? DateCode(culture) + " " + time : hasDate ? DateCode(culture) : time;
    }

    private static bool HasMeridiem(string text, DateTimeFormatInfo format)
    {
        foreach (var designator in new[] { format.AMDesignator, format.PMDesignator, "AM", "PM", "A", "P" })
        {
            if (designator.Length > 0 && text.EndsWith(designator, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    // A date part is a date separator, a dash or slash, or a word other than AM or PM.
    private static bool HasDatePart(string text, DateTimeFormatInfo format)
    {
        if (text.AsSpan().IndexOfAny('/', '-') >= 0 || (format.DateSeparator.Length > 0 && text.Contains(format.DateSeparator, StringComparison.Ordinal)))
            return true;
        foreach (var word in text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (char.IsLetter(word[0]) && !HasMeridiem(word, format))
                return true;
        }

        return false;
    }

    // The culture's short date as a format code: English Excel's m/d/yyyy for the invariant culture.
    internal static string DateCode(CultureInfo culture)
    {
        if (culture.Name.Length == 0)
            return "m/d/yyyy";
        var pattern = culture.DateTimeFormat.ShortDatePattern;
        var code = new System.Text.StringBuilder(pattern.Length);
        foreach (var c in pattern)
        {
            // Other separators, such as the dot of dd.MM.yyyy, are literals in a format code.
            if (c == 'M')
                code.Append('m');
            else if (c is 'd' or 'y' or '/' or '-' or ' ')
                code.Append(c);
            else
                code.Append('\\').Append(c);
        }

        return code.ToString();
    }

    internal static string TimeCode(CultureInfo culture, bool seconds)
    {
        var twelveHour = culture.Name.Length == 0 || culture.DateTimeFormat.LongTimePattern.Contains('t', StringComparison.Ordinal);
        return "h:mm" + (seconds ? ":ss" : "") + (twelveHour ? " AM/PM" : "");
    }
}
