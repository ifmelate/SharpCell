using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// Text functions. Lengths and positions count UTF-16 code units, as Excel does (an emoji is two
/// characters). Results longer than Excel's 32 767 characters are <c>#VALUE!</c>. The byte
/// variants (LENB, LEFTB...) behave like the plain ones, as in Excel outside double-byte locales.
/// </summary>
internal static class TextFunctions
{
    private static readonly ArgumentKind[] Values = [ArgumentKind.Value];
    private static readonly ArgumentKind[] AnyArguments = [ArgumentKind.Any];

    // CHAR and CODE use Windows-1252, the code page of Excel in Western locales. Bytes 128-159
    // differ from Latin-1; the five bytes 1252 leaves undefined map to the same control characters.
    private static readonly char[] Windows1252High =
    [
        '€', '\u0081', '‚', 'ƒ', '„', '…', '†', '‡',
        'ˆ', '‰', 'Š', '‹', 'Œ', '\u008D', 'Ž', '\u008F',
        '\u0090', '‘', '’', '“', '”', '•', '–', '—',
        '˜', '™', 'š', '›', 'œ', '\u009D', 'ž', 'Ÿ',
    ];

    public static void Register(FunctionRegistry registry)
    {
        var max = FunctionRegistry.MaxArguments;
        Add(registry, 1, 1, Len, "LEN", "LENB");
        Add(registry, 1, 2, Left, "LEFT", "LEFTB");
        Add(registry, 1, 2, Right, "RIGHT", "RIGHTB");
        Add(registry, 3, 3, Mid, "MID", "MIDB");
        Add(registry, 4, 4, Replace, "REPLACE", "REPLACEB");
        Add(registry, 2, 3, Find, "FIND", "FINDB");
        Add(registry, 2, 3, Search, "SEARCH", "SEARCHB");
        Add(registry, 3, 4, Substitute, "SUBSTITUTE");
        Add(registry, 2, 2, Rept, "REPT");
        Add(registry, 1, 1, call => MapText(call, (text, culture) => culture.TextInfo.ToUpper(text)), "UPPER");
        Add(registry, 1, 1, call => MapText(call, (text, culture) => culture.TextInfo.ToLower(text)), "LOWER");
        Add(registry, 1, 1, call => MapText(call, Proper), "PROPER");
        Add(registry, 1, 1, call => MapText(call, (text, _) => Trim(text)), "TRIM");
        Add(registry, 1, 1, call => MapText(call, (text, _) => Clean(text)), "CLEAN");
        Add(registry, 2, 2, Exact, "EXACT");
        Add(registry, 1, 1, Char, "CHAR");
        Add(registry, 1, 1, Code, "CODE");
        Add(registry, 1, 1, UniChar, "UNICHAR");
        Add(registry, 1, 1, UniCode, "UNICODE");
        Add(registry, 1, 1, T, "T");
        Add(registry, 1, 1, Value, "VALUE");
        Add(registry, 1, 3, NumberValue, "NUMBERVALUE");
        Add(registry, 1, 2, ValueToText, "VALUETOTEXT");
        registry.Add(new FunctionInfo("TEXT", 2, 2, Values, Text)
        {
            Status = FunctionStatus.KnownDeviation,
            Deviation = "Calendar and era codes (B1, B2, e, g, [DBNum]) are not interpreted, and month names and separators come from the workbook culture, not from [$-xxxx] locale tags.",
        });
        Add(registry, 1, 3, Fixed, "FIXED");
        Add(registry, 1, 2, Dollar, "DOLLAR");
        Add(registry, 1, max, Concatenate, "CONCATENATE");
        registry.Add(new FunctionInfo("CONCAT", 1, max, AnyArguments, Concat));
        registry.Add(new FunctionInfo("TEXTJOIN", 3, max, [ArgumentKind.Any, ArgumentKind.Value, ArgumentKind.Any], TextJoin));
        registry.Add(new FunctionInfo("ARRAYTOTEXT", 1, 2, [ArgumentKind.Any, ArgumentKind.Value], ArrayToText));
    }

    private static void Add(FunctionRegistry registry, int min, int max, FunctionBody body, params string[] names)
    {
        foreach (var name in names)
            registry.Add(new FunctionInfo(name, min, max, Values, body));
    }

    private static CellValue Error(ErrorKind kind) => CellValue.Error(kind);

    /// <summary>Text no longer than Excel allows, else <c>#VALUE!</c>.</summary>
    internal static CellValue Result(string text) => text.Length > Operators.MaxTextLength ? Error(ErrorKind.Value) : CellValue.Text(text);

    // An optional count read as an integer: absent gives the default, left empty (LEFT(A1,)) gives 0.
    private static CellValue OptionalInteger(FunctionCall call, int index, double absent) =>
        index < call.Count ? call.Integer(index) : CellValue.Number(absent);

    private static Operand MapText(FunctionCall call, Func<string, CultureInfo, string> map)
    {
        var text = call.Text(0);
        return text.IsError ? text : CellValue.Text(map(text.AsText(), call.Context.Culture));
    }

    private static Operand Len(FunctionCall call)
    {
        var text = call.Text(0);
        return text.IsError ? text : CellValue.Number(text.AsText().Length);
    }

    private static Operand Left(FunctionCall call) => Take(call, fromEnd: false);

    private static Operand Right(FunctionCall call) => Take(call, fromEnd: true);

    private static Operand Take(FunctionCall call, bool fromEnd)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        var count = OptionalInteger(call, 1, 1);
        if (count.IsError)
            return count;

        var n = count.AsNumber();
        if (n < 0)
            return Error(ErrorKind.Value);
        var s = text.AsText();
        if (n >= s.Length)
            return text;
        return CellValue.Text(fromEnd ? s[^(int)n..] : s[..(int)n]);
    }

    // MID(text, start, count): start before 1 or a negative count is #VALUE!; past the end is "".
    private static Operand Mid(FunctionCall call)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        var start = call.Integer(1);
        if (start.IsError)
            return start;
        var count = call.Integer(2);
        if (count.IsError)
            return count;

        double from = start.AsNumber(), n = count.AsNumber();
        if (from < 1 || n < 0)
            return Error(ErrorKind.Value);
        var s = text.AsText();
        if (from > s.Length)
            return CellValue.Text("");
        var index = (int)from - 1;
        return CellValue.Text(s.Substring(index, (int)Math.Min(n, s.Length - index)));
    }

    // REPLACE(old, start, count, new): a start past the end appends.
    private static Operand Replace(FunctionCall call)
    {
        var old = call.Text(0);
        if (old.IsError)
            return old;
        var start = call.Integer(1);
        if (start.IsError)
            return start;
        var count = call.Integer(2);
        if (count.IsError)
            return count;
        var replacement = call.Text(3);
        if (replacement.IsError)
            return replacement;

        double from = start.AsNumber(), n = count.AsNumber();
        if (from < 1 || n < 0)
            return Error(ErrorKind.Value);
        var s = old.AsText();
        var index = (int)Math.Min(from - 1, s.Length);
        var end = (int)Math.Min(index + n, s.Length);
        return Result(string.Concat(s.AsSpan(0, index), replacement.AsText(), s.AsSpan(end)));
    }

    private static Operand Find(FunctionCall call) => Locate(call, (find, within, start, _) =>
        within.IndexOf(find, start, StringComparison.Ordinal));

    private static Operand Search(FunctionCall call) => Locate(call, (find, within, start, culture) =>
        TextMatch.SearchIndex(find, within, start, culture, call.Context));

    // FIND and SEARCH: the 1-based position of the first match at or after start (default 1).
    // Empty text to find matches at start; start must lie within the text or just after it.
    private static Operand Locate(FunctionCall call, Func<string, string, int, CultureInfo, int> indexOf)
    {
        var find = call.Text(0);
        if (find.IsError)
            return find;
        var within = call.Text(1);
        if (within.IsError)
            return within;
        var start = OptionalInteger(call, 2, 1);
        if (start.IsError)
            return start;

        var s = within.AsText();
        var from = start.AsNumber();
        if (from < 1 || from > s.Length + 1)
            return Error(ErrorKind.Value);
        var index = indexOf(find.AsText(), s, (int)from - 1, call.Context.Culture);
        return index < 0 ? Error(ErrorKind.Value) : CellValue.Number(index + 1);
    }

    // SUBSTITUTE(text, old, new, [instance]): every occurrence, or only the instance-th one.
    private static Operand Substitute(FunctionCall call)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        var old = call.Text(1);
        if (old.IsError)
            return old;
        var replacement = call.Text(2);
        if (replacement.IsError)
            return replacement;
        var instance = 0.0;
        if (call.Count > 3)
        {
            var n = call.Integer(3);
            if (n.IsError)
                return n;
            instance = n.AsNumber();
            if (instance < 1)
                return Error(ErrorKind.Value);
        }

        string s = text.AsText(), find = old.AsText();
        if (find.Length == 0)
            return text;
        if (instance == 0)
            return Result(s.Replace(find, replacement.AsText(), StringComparison.Ordinal));

        var index = -1;
        for (var i = 0; i < instance; i++)
        {
            index = s.IndexOf(find, index < 0 ? 0 : index + find.Length, StringComparison.Ordinal);
            if (index < 0)
                return text;
        }

        return Result(string.Concat(s.AsSpan(0, index), replacement.AsText(), s.AsSpan(index + find.Length)));
    }

    private static Operand Rept(FunctionCall call)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        var count = call.Integer(1);
        if (count.IsError)
            return count;

        var n = count.AsNumber();
        var s = text.AsText();
        if (n < 0 || s.Length * n > Operators.MaxTextLength)
            return Error(ErrorKind.Value);
        if (s.Length == 0)
            return text;
        var sb = new StringBuilder(s.Length * (int)n);
        for (var i = 0; i < (int)n; i++)
            sb.Append(s);
        return CellValue.Text(sb.ToString());
    }

    // A letter after anything but a letter starts a word: "2nd" becomes "2Nd", as in Excel.
    private static string Proper(string text, CultureInfo culture)
    {
        var chars = text.ToCharArray();
        var previousLetter = false;
        for (var i = 0; i < chars.Length; i++)
        {
            var letter = char.IsLetter(chars[i]);
            if (letter)
                chars[i] = previousLetter ? culture.TextInfo.ToLower(chars[i]) : culture.TextInfo.ToUpper(chars[i]);
            previousLetter = letter;
        }

        return new string(chars);
    }

    // Removes leading and trailing spaces and collapses inner runs to one. Only the ASCII space
    // counts; tabs and no-break spaces stay.
    private static string Trim(string text)
    {
        var sb = new StringBuilder(text.Length);
        foreach (var word in text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (sb.Length > 0)
                sb.Append(' ');
            sb.Append(word);
        }

        return sb.ToString();
    }

    // Removes the ASCII control characters 0-31.
    private static string Clean(string text)
    {
        var sb = new StringBuilder(text.Length);
        foreach (var c in text)
        {
            if (c >= ' ')
                sb.Append(c);
        }

        return sb.ToString();
    }

    private static Operand Exact(FunctionCall call)
    {
        var a = call.Text(0);
        if (a.IsError)
            return a;
        var b = call.Text(1);
        if (b.IsError)
            return b;
        return CellValue.Boolean(string.Equals(a.AsText(), b.AsText(), StringComparison.Ordinal));
    }

    private static Operand Char(FunctionCall call)
    {
        var code = call.Integer(0);
        if (code.IsError)
            return code;
        var n = code.AsNumber();
        if (n is < 1 or > 255)
            return Error(ErrorKind.Value);
        var c = n is >= 128 and < 160 ? Windows1252High[(int)n - 128] : (char)n;
        return CellValue.Text(c.ToString());
    }

    // The Windows-1252 code of the first character; characters the code page lacks give 63 ('?').
    private static Operand Code(FunctionCall call)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        var s = text.AsText();
        if (s.Length == 0)
            return Error(ErrorKind.Value);

        var c = s[0];
        if (c < 128 || (c >= 160 && c <= 255))
            return CellValue.Number(c);
        var index = Array.IndexOf(Windows1252High, c);
        return CellValue.Number(index >= 0 ? 128 + index : 63);
    }

    // UNICHAR(number): 0 or beyond U+10FFFF is #VALUE!; a lone surrogate is #N/A.
    private static Operand UniChar(FunctionCall call)
    {
        var code = call.Integer(0);
        if (code.IsError)
            return code;
        var n = code.AsNumber();
        if (n is < 1 or > 0x10FFFF)
            return Error(ErrorKind.Value);
        if (n is >= 0xD800 and <= 0xDFFF)
            return Error(ErrorKind.NA);
        return CellValue.Text(char.ConvertFromUtf32((int)n));
    }

    private static Operand UniCode(FunctionCall call)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        var s = text.AsText();
        if (s.Length == 0)
            return Error(ErrorKind.Value);
        return CellValue.Number(char.IsSurrogatePair(s, 0) ? char.ConvertToUtf32(s[0], s[1]) : s[0]);
    }

    private static Operand T(FunctionCall call)
    {
        var value = call.Value(0);
        return value.Kind is CellValueKind.Text or CellValueKind.Error ? value : CellValue.Text("");
    }

    // VALUE reads text as a typed entry would be read; TRUE and FALSE are not numbers here.
    private static Operand Value(FunctionCall call)
    {
        var value = call.Value(0);
        if (value.Kind == CellValueKind.Boolean)
            return Error(ErrorKind.Value);
        return Coercion.ToNumber(value, call.Context.Culture, call.Context.DateSystem);
    }

    // NUMBERVALUE(text, [decimal], [group]): the separators are given (the first character of
    // each; the culture's by default). Spaces anywhere are ignored, group separators only before
    // the decimal separator, and each trailing % divides by 100.
    private static Operand NumberValue(FunctionCall call)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        var format = call.Context.Culture.NumberFormat;
        var decimalText = call.Text(1, format.NumberDecimalSeparator);
        if (decimalText.IsError)
            return decimalText;
        var groupText = call.Text(2, format.NumberGroupSeparator);
        if (groupText.IsError)
            return groupText;

        var decimalSeparator = decimalText.AsText().Length > 0 ? decimalText.AsText()[0] : format.NumberDecimalSeparator[0];
        var groupSeparator = groupText.AsText().Length > 0 ? groupText.AsText()[0] : (char?)null;
        if (groupSeparator == decimalSeparator)
            return Error(ErrorKind.Value);

        var s = text.AsText();
        var sb = new StringBuilder(s.Length);
        var percent = 0;
        var seenDecimal = false;
        foreach (var c in s)
        {
            if (char.IsWhiteSpace(c))
                continue;
            if (c == '%')
            {
                percent++;
                continue;
            }

            // Nothing but spaces and percent signs may follow a percent sign.
            if (percent > 0)
                return Error(ErrorKind.Value);
            if (c == decimalSeparator)
            {
                if (seenDecimal)
                    return Error(ErrorKind.Value);
                seenDecimal = true;
                sb.Append('.');
            }
            else if (c == groupSeparator)
            {
                if (seenDecimal)
                    return Error(ErrorKind.Value);
            }
            else
            {
                sb.Append(c);
            }
        }

        if (sb.Length == 0)
            return CellValue.Number(0);
        if (!double.TryParse(sb.ToString(), NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent,
                CultureInfo.InvariantCulture, out var number))
            return Error(ErrorKind.Value);
        return CellValue.Number(number / Math.Pow(100, percent));
    }

    // VALUETOTEXT(value, [format]): 0 gives the value as a cell shows it, 1 as a formula would
    // write it (text in quotes). Errors become their text.
    private static Operand ValueToText(FunctionCall call)
    {
        var format = OptionalInteger(call, 1, 0);
        if (format.IsError)
            return format;
        if (format.AsNumber() is not (0 or 1))
            return Error(ErrorKind.Value);
        return ToDisplayText(call.Value(0), strict: format.AsNumber() == 1, call.Context.Culture);
    }

    private static CellValue ToDisplayText(CellValue value, bool strict, CultureInfo culture)
    {
        switch (value.Kind)
        {
            case CellValueKind.Error:
                return CellValue.Text(value.AsError().ToText());
            case CellValueKind.Text when strict:
                return Result("\"" + value.AsText().Replace("\"", "\"\"", StringComparison.Ordinal) + "\"");
            default:
                var text = Coercion.ToText(value, culture);
                return text.IsError ? CellValue.Text(ErrorKind.Value.ToText()) : text;
        }
    }

    // TEXT(value, format): numbers (and text that reads as one) through the format's number
    // sections, other text through its text section.
    private static Operand Text(FunctionCall call)
    {
        var value = call.Value(0);
        if (value.IsError)
            return value;
        var code = call.Text(1);
        if (code.IsError)
            return code;
        if (!NumberFormat.TryParse(code.AsText(), out var format))
            return Error(ErrorKind.Value);

        var context = call.Context;
        double number;
        switch (value.Kind)
        {
            case CellValueKind.Number:
                number = value.AsNumber();
                break;
            case CellValueKind.Text:
                var converted = Coercion.ToNumber(value, context.Culture, context.DateSystem);
                if (converted.IsError)
                    return Result(format!.FormatText(value.AsText()));
                number = converted.AsNumber();
                break;
            case CellValueKind.Boolean:
                return Result(format!.FormatText(value.AsBoolean() ? "TRUE" : "FALSE"));
            default:
                number = 0;
                break;
        }

        var text = format!.Format(number, context.Culture, context.DateSystem);
        return text is null ? Error(ErrorKind.Value) : Result(text);
    }

    // FIXED(number, [decimals], [no_commas]): rounded to the decimals (negative ones round left
    // of the point) and written with group separators unless no_commas.
    private static Operand Fixed(FunctionCall call)
    {
        var number = call.Number(0);
        if (number.IsError)
            return number;
        var decimals = OptionalInteger(call, 1, 2);
        if (decimals.IsError)
            return decimals;
        var noCommas = call.Boolean(2, false);
        if (noCommas.IsError)
            return noCommas;

        var places = decimals.AsNumber();
        if (places > 127)
            return Error(ErrorKind.Value);
        var code = (noCommas.AsBoolean() ? "0" : "#,##0") + Decimals(places);
        return FormatRounded(call, number.AsNumber(), places, code);
    }

    // DOLLAR(number, [decimals]): currency text in the culture's style; without a culture as
    // English Excel writes it, $1,234.57 and ($1,234.57).
    private static Operand Dollar(FunctionCall call)
    {
        var number = call.Number(0);
        if (number.IsError)
            return number;
        var decimals = OptionalInteger(call, 1, 2);
        if (decimals.IsError)
            return decimals;

        var places = decimals.AsNumber();
        if (places > 127)
            return Error(ErrorKind.Value);
        var amount = "#,##0" + Decimals(places);
        var culture = call.Context.Culture;
        string code;
        if (culture.Name.Length == 0)
        {
            code = "$" + amount + ";($" + amount + ")";
        }
        else
        {
            var format = culture.NumberFormat;
            var symbol = "\"" + format.CurrencySymbol.Replace("\"", "", StringComparison.Ordinal) + "\"";
            var positive = format.CurrencyPositivePattern switch
            {
                0 => symbol + amount,
                1 => amount + symbol,
                2 => symbol + " " + amount,
                _ => amount + " " + symbol,
            };
            var negative = format.CurrencyNegativePattern switch
            {
                0 => "(" + symbol + amount + ")",
                1 => "-" + symbol + amount,
                2 => symbol + "-" + amount,
                3 => symbol + amount + "-",
                4 => "(" + amount + symbol + ")",
                5 => "-" + amount + symbol,
                6 => amount + "-" + symbol,
                7 => amount + symbol + "-",
                8 => "-" + amount + " " + symbol,
                9 => "-" + symbol + " " + amount,
                10 => amount + " " + symbol + "-",
                11 => symbol + " " + amount + "-",
                12 => symbol + " -" + amount,
                13 => amount + "- " + symbol,
                14 => "(" + symbol + " " + amount + ")",
                15 => "(" + amount + " " + symbol + ")",
                _ => symbol + "- " + amount,
            };
            code = positive + ";" + negative;
        }

        return FormatRounded(call, number.AsNumber(), places, code);
    }

    private static string Decimals(double places) => places > 0 ? "." + new string('0', (int)places) : "";

    private static CellValue FormatRounded(FunctionCall call, double number, double places, string code)
    {
        if (places < 0)
        {
            var factor = Math.Pow(10, Math.Min(-places, 308));
            number = Math.Round(number / factor, MidpointRounding.AwayFromZero) * factor;
            if (!double.IsFinite(number))
                return Error(ErrorKind.Num);
        }

        NumberFormat.TryParse(code, out var format);
        var text = format!.Format(number, call.Context.Culture, call.Context.DateSystem);
        return text is null ? Error(ErrorKind.Value) : Result(text);
    }

    private static Operand Concatenate(FunctionCall call)
    {
        var sb = new StringBuilder();
        for (var i = 0; i < call.Count; i++)
        {
            var text = call.Text(i);
            if (text.IsError)
                return text;
            sb.Append(text.AsText());
            if (sb.Length > Operators.MaxTextLength)
                return Error(ErrorKind.Value);
        }

        return CellValue.Text(sb.ToString());
    }

    // CONCAT joins every value of its arguments, ranges row by row; the first error wins.
    private static Operand Concat(FunctionCall call)
    {
        var sb = new StringBuilder();
        CellValue? error = null;
        var culture = call.Context.Culture;
        TextMatch.ForEachItem(call, 0, dense: false, value =>
        {
            var text = Coercion.ToText(value, culture);
            if (text.IsError)
            {
                error = text;
                return false;
            }

            sb.Append(text.AsText());
            if (sb.Length <= Operators.MaxTextLength)
                return true;
            error = Error(ErrorKind.Value);
            return false;
        });
        return error ?? CellValue.Text(sb.ToString());
    }

    // TEXTJOIN(delimiter, ignore_empty, text1, ...): a range or array of delimiters is used in
    // turn. Without ignore_empty, empty cells count as empty texts between delimiters.
    private static Operand TextJoin(FunctionCall call)
    {
        var context = call.Context;
        var delimiterValue = Evaluator.ToValue(call[0], context);
        var delimiters = new List<string>();
        foreach (var item in delimiterValue.Kind == CellValueKind.Array ? delimiterValue.AsArray() : new[,] { { delimiterValue } })
        {
            var text = Coercion.ToText(item, context.Culture);
            if (text.IsError)
                return text;
            delimiters.Add(text.AsText());
        }

        var ignoreEmpty = call.Boolean(1);
        if (ignoreEmpty.IsError)
            return ignoreEmpty;

        var skipEmpty = ignoreEmpty.AsBoolean();
        var sb = new StringBuilder();
        CellValue? error = null;
        var count = 0;
        TextMatch.ForEachItem(call, 2, dense: !skipEmpty && delimiters.Exists(d => d.Length > 0), value =>
        {
            var text = Coercion.ToText(value, context.Culture);
            if (text.IsError)
            {
                error = text;
                return false;
            }

            if (skipEmpty && text.AsText().Length == 0)
                return true;
            if (count > 0)
                sb.Append(delimiters[(count - 1) % delimiters.Count]);
            sb.Append(text.AsText());
            count++;
            if (sb.Length <= Operators.MaxTextLength)
                return true;
            error = Error(ErrorKind.Value);
            return false;
        });
        return error ?? CellValue.Text(sb.ToString());
    }

    // ARRAYTOTEXT(array, [format]): 0 lists the values separated by ", "; 1 writes an array
    // constant, {1,"a";TRUE,#N/A}.
    private static Operand ArrayToText(FunctionCall call)
    {
        var format = OptionalInteger(call, 1, 0);
        if (format.IsError)
            return format;
        if (format.AsNumber() is not (0 or 1))
            return Error(ErrorKind.Value);

        var strict = format.AsNumber() == 1;
        var value = Evaluator.ToValue(call[0], call.Context);
        var array = value.Kind == CellValueKind.Array ? value.AsArray() : new[,] { { value } };
        var sb = new StringBuilder();
        if (strict)
            sb.Append('{');
        for (var r = 0; r < array.GetLength(0); r++)
        {
            call.Context.CancellationToken.ThrowIfCancellationRequested();
            for (var c = 0; c < array.GetLength(1); c++)
            {
                if (r > 0 || c > 0)
                    sb.Append(!strict ? ", " : c == 0 ? ";" : ",");
                sb.Append(ToDisplayText(array[r, c], strict, call.Context.Culture) is { Kind: CellValueKind.Text } text ? text.AsText() : "");
                if (sb.Length > Operators.MaxTextLength)
                    return Error(ErrorKind.Value);
            }
        }

        if (strict)
            sb.Append('}');
        return Result(sb.ToString());
    }
}
