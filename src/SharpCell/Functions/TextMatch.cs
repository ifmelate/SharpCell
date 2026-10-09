using System;
using System.Collections.Generic;
using System.Globalization;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>Searching inside text and walking the values of text functions' range arguments.</summary>
internal static class TextMatch
{
    private const int CancellationCheckInterval = 4096;

    /// <summary>
    /// SEARCH: the index of the first match at or after <paramref name="start"/>, ignoring case,
    /// with Excel wildcards (<c>*</c>, <c>?</c>, <c>~</c> escapes); -1 if none.
    /// </summary>
    public static int SearchIndex(string find, string within, int start, CultureInfo culture, EvaluationContext context)
    {
        // Upper-casing keeps the length (TextInfo maps character by character), so indexes carry over.
        var pattern = culture.TextInfo.ToUpper(find);
        var text = culture.TextInfo.ToUpper(within);
        if (pattern.AsSpan().IndexOfAny("*?~") < 0)
            return text.IndexOf(pattern, start, StringComparison.Ordinal);

        var tokens = Parse(pattern);
        for (var i = start; i <= text.Length; i++)
        {
            if ((i - start) % CancellationCheckInterval == CancellationCheckInterval - 1)
                context.CancellationToken.ThrowIfCancellationRequested();
            if (MatchesAt(tokens, text, i))
                return i;
        }

        return -1;
    }

    /// <summary>
    /// The index of <paramref name="find"/> in <paramref name="text"/> at or after start, by
    /// ordinal comparison or ignoring case; -1 if none.
    /// </summary>
    public static int IndexOf(string text, string find, int start, bool ignoreCase, CultureInfo culture)
    {
        if (!ignoreCase)
            return text.IndexOf(find, start, StringComparison.Ordinal);
        return culture.TextInfo.ToUpper(text).IndexOf(culture.TextInfo.ToUpper(find), start, StringComparison.Ordinal);
    }

    // A pattern position: a character, '?' (any one character) or '*' (any run).
    private readonly record struct Token(char Char, bool AnyOne, bool AnyRun);

    private static List<Token> Parse(string pattern)
    {
        var tokens = new List<Token>(pattern.Length);
        for (var i = 0; i < pattern.Length; i++)
        {
            var c = pattern[i];
            if (c == '~')
            {
                // A tilde makes the next character literal, whatever it is; a final one is dropped.
                if (i + 1 < pattern.Length)
                    tokens.Add(new Token(pattern[++i], false, false));
            }
            else if (c == '*')
                tokens.Add(new Token(c, false, true));
            else if (c == '?')
                tokens.Add(new Token(c, true, false));
            else
                tokens.Add(new Token(c, false, false));
        }

        return tokens;
    }

    // Whether the pattern matches some text starting at index: greedy, with one backtrack
    // point per '*'. The pattern only has to match a prefix.
    private static bool MatchesAt(List<Token> tokens, string text, int index)
    {
        int pi = 0, ti = index, starPattern = -1, starText = 0;
        while (true)
        {
            if (pi == tokens.Count)
                return true;
            var token = tokens[pi];
            if (token.AnyRun)
            {
                starPattern = ++pi;
                starText = ti;
                continue;
            }

            if (ti < text.Length && (token.AnyOne || token.Char == text[ti]))
            {
                pi++;
                ti++;
                continue;
            }

            if (starPattern < 0 || starText >= text.Length)
                return false;
            pi = starPattern;
            ti = ++starText;
        }
    }

    /// <summary>
    /// Visits the values of arguments <paramref name="first"/> on: each cell of a reference row by
    /// row (only cells that exist unless <paramref name="dense"/>), each element of an array, a
    /// direct value as it is (an omitted argument as an empty value). The visitor returns false to
    /// stop. Once a dirty cell has been read, the remaining referenced cells are still read so that
    /// all dirty inputs are found in one pass.
    /// </summary>
    public static void ForEachItem(FunctionCall call, int first, bool dense, Func<CellValue, bool> visit)
    {
        var context = call.Context;
        var visiting = true;
        var visited = 0;
        for (var i = first; i < call.Count; i++)
        {
            var argument = call[i];
            if (argument.Reference is { } reference)
            {
                foreach (var (sheet, area) in reference.Areas)
                {
                    if (visiting && dense)
                    {
                        for (var row = area.FirstRow; row <= area.LastRow && visiting; row++)
                        {
                            for (var column = area.FirstColumn; column <= area.LastColumn && visiting; column++)
                            {
                                if (++visited % CancellationCheckInterval == 0)
                                    context.CancellationToken.ThrowIfCancellationRequested();
                                visiting = visit(context.ReadCell(sheet, row, column));
                            }
                        }

                        if (visiting)
                            continue;
                        if (!call.MetPendingInput)
                            return;
                    }

                    foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
                    {
                        if (++visited % CancellationCheckInterval == 0)
                            context.CancellationToken.ThrowIfCancellationRequested();
                        var value = context.ReadCell(sheet, cell);
                        if (visiting && !visit(value))
                            visiting = false;
                        if (!visiting && !call.MetPendingInput)
                            return;
                    }
                }
            }
            else if (!visiting)
            {
                continue;
            }
            else if (argument.Value.Kind == CellValueKind.Array)
            {
                foreach (var element in argument.Value.AsArray())
                {
                    if (!visit(element))
                    {
                        visiting = false;
                        break;
                    }
                }
            }
            else
            {
                visiting = visit(argument.Value.Kind == CellValueKind.Missing ? CellValue.Empty : argument.Value);
            }

            if (!visiting && !call.MetPendingInput)
                return;
        }
    }
}
