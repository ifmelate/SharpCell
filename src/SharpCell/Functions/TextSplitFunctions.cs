using System;
using System.Collections.Generic;
using System.Globalization;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// TEXTSPLIT, TEXTBEFORE and TEXTAFTER. Delimiters may be a range or array of several texts;
/// at each position the first of them that matches is taken.
/// </summary>
internal static class TextSplitFunctions
{
    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("TEXTSPLIT", 2, 6,
            [ArgumentKind.Value, ArgumentKind.Any, ArgumentKind.Any, ArgumentKind.Value], TextSplit));
        registry.Add(new FunctionInfo("TEXTBEFORE", 2, 6, [ArgumentKind.Value, ArgumentKind.Any, ArgumentKind.Value], call => Around(call, before: true)));
        registry.Add(new FunctionInfo("TEXTAFTER", 2, 6, [ArgumentKind.Value, ArgumentKind.Any, ArgumentKind.Value], call => Around(call, before: false)));
    }

    private static CellValue Error(ErrorKind kind) => CellValue.Error(kind);

    // The delimiters of argument index: null when it is absent or left empty, an error value
    // when one of them is an error or empty text.
    private static bool TryDelimiters(FunctionCall call, int index, out List<string>? delimiters, out CellValue error)
    {
        delimiters = null;
        error = default;
        if (!call.Has(index))
            return true;

        var value = Evaluator.ToValue(call[index], call.Context);
        delimiters = [];
        foreach (var item in value.Kind == CellValueKind.Array ? value.AsArray() : new[,] { { value } })
        {
            var text = Coercion.ToText(item, call.Context.Culture);
            if (text.IsError)
            {
                error = text;
                return false;
            }

            delimiters.Add(text.AsText());
        }

        return true;
    }

    // A flag argument that must be 0 or 1 (TRUE and FALSE included); absent or empty is 0.
    private static CellValue Flag(FunctionCall call, int index)
    {
        var flag = call.Integer(index, 0);
        if (flag.IsError)
            return flag;
        return flag.AsNumber() is 0 or 1 ? flag : Error(ErrorKind.Value);
    }

    // The first delimiter occurrence at or after start: its index and length; -1 if none.
    private static (int Index, int Length) NextDelimiter(string text, List<string> delimiters, int start, bool ignoreCase, CultureInfo culture)
    {
        var best = -1;
        var length = 0;
        foreach (var delimiter in delimiters)
        {
            var index = TextMatch.IndexOf(text, delimiter, start, ignoreCase, culture);
            if (index >= 0 && (best < 0 || index < best))
            {
                best = index;
                length = delimiter.Length;
            }
        }

        return (best, length);
    }

    // TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]):
    // rows split by the row delimiters, each row by the column delimiters; short rows are padded
    // (with #N/A by default). Empty delimiters, or none at all, are #VALUE!.
    private static Operand TextSplit(FunctionCall call)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        if (!TryDelimiters(call, 1, out var columns, out var error) || !TryDelimiters(call, 2, out var rows, out error))
            return error;
        var ignoreEmpty = call.Boolean(3, false);
        if (ignoreEmpty.IsError)
            return ignoreEmpty;
        var mode = Flag(call, 4);
        if (mode.IsError)
            return mode;
        var pad = call.Has(5) ? call.Value(5) : Error(ErrorKind.NA);

        if ((columns is null && rows is null) || columns?.Exists(d => d.Length == 0) == true || rows?.Exists(d => d.Length == 0) == true)
            return Error(ErrorKind.Value);

        var ignoreCase = mode.AsNumber() == 1;
        var skip = ignoreEmpty.AsBoolean();
        var culture = call.Context.Culture;
        var lines = Split(text.AsText(), rows, skip, ignoreCase, culture);
        var table = new List<List<string>>(lines.Count);
        var width = 1;
        foreach (var line in lines)
        {
            call.Context.CancellationToken.ThrowIfCancellationRequested();
            var cells = Split(line, columns, skip, ignoreCase, culture);
            if (cells.Count == 0)
                continue;
            table.Add(cells);
            width = Math.Max(width, cells.Count);
        }

        if (table.Count == 0)
            return Error(ErrorKind.Calc);
        if ((long)table.Count * width > Evaluator.MaxArrayCells)
            return Error(ErrorKind.Num);

        var result = new CellValue[table.Count, width];
        for (var r = 0; r < table.Count; r++)
        {
            for (var c = 0; c < width; c++)
                result[r, c] = c < table[r].Count ? CellValue.Text(table[r][c]) : pad;
        }

        return result.Length == 1 ? result[0, 0] : CellValue.Array(result);
    }

    private static List<string> Split(string text, List<string>? delimiters, bool skipEmpty, bool ignoreCase, CultureInfo culture)
    {
        var parts = new List<string>();
        if (delimiters is null)
        {
            parts.Add(text);
            return parts;
        }

        var start = 0;
        while (true)
        {
            var (index, length) = NextDelimiter(text, delimiters, start, ignoreCase, culture);
            var part = index < 0 ? text[start..] : text[start..index];
            if (!skipEmpty || part.Length > 0)
                parts.Add(part);
            if (index < 0)
                return parts;
            start = index + length;
        }
    }

    // The if_not_found value; an empty cell gives empty text, not 0.
    private static CellValue NotFound(CellValue value) => value.Kind == CellValueKind.Empty ? CellValue.Text("") : value;

    // TEXTBEFORE and TEXTAFTER(text, delimiter, [instance], [match_mode], [match_end], [if_not_found]).
    // A negative instance counts from the end. With match_end the end of the text (the start, for
    // a negative instance) also counts as a delimiter. An empty delimiter matches at once.
    private static Operand Around(FunctionCall call, bool before)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        if (!TryDelimiters(call, 1, out var delimiters, out var error))
            return error;
        delimiters ??= [""];
        // The instance is rounded down, not truncated: -1.2 is -2.
        var instance = call.Number(2, 1);
        if (instance.IsError)
            return instance;
        var mode = Flag(call, 3);
        if (mode.IsError)
            return mode;
        var matchEnd = Flag(call, 4);
        if (matchEnd.IsError)
            return matchEnd;

        var s = text.AsText();
        var n = Math.Floor(instance.AsNumber());
        var shortest = int.MaxValue;
        foreach (var delimiter in delimiters)
            shortest = Math.Min(shortest, delimiter.Length);

        // An instance that cannot fit in the text is #VALUE!; Excel checks this only when the
        // instance is given.
        if (n == 0 || (call.Has(2) && Math.Abs(n) > s.Length - shortest + 1))
            return Error(ErrorKind.Value);

        var ignoreCase = mode.AsNumber() == 1;
        var culture = call.Context.Culture;
        var occurrences = new List<(int Index, int Length)>();
        if (shortest == 0)
        {
            // An empty delimiter matches at the start, or at the end when counting from it.
            occurrences.Add(n > 0 ? (0, 0) : (s.Length, 0));
        }
        else
        {
            var start = 0;
            while (start <= s.Length)
            {
                var (index, length) = NextDelimiter(s, delimiters, start, ignoreCase, culture);
                if (index < 0)
                    break;
                occurrences.Add((index, length));
                start = index + length;
            }
        }

        if (matchEnd.AsNumber() == 1 && shortest > 0)
        {
            if (n > 0)
                occurrences.Add((s.Length, 0));
            else
                occurrences.Insert(0, (0, 0));
        }

        var position = n > 0 ? (int)Math.Min(n, int.MaxValue) - 1 : occurrences.Count + (int)Math.Max(n, int.MinValue + 1);
        if (shortest == 0)
            position = 0;
        if (position < 0 || position >= occurrences.Count)
            return call.Has(5) ? NotFound(call.Value(5)) : Error(ErrorKind.NA);

        var (at, size) = occurrences[position];
        return CellValue.Text(before ? s[..at] : s[(at + size)..]);
    }
}
