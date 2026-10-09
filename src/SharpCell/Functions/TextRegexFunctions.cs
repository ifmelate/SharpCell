using System;
using System.Text.RegularExpressions;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// REGEXTEST, REGEXEXTRACT and REGEXREPLACE on .NET regular expressions. Excel uses PCRE2; the
/// common syntax is the same, the corners are not (see the deviation note). A pattern that does
/// not parse or a match that runs past the timeout is <c>#VALUE!</c>.
/// </summary>
internal static class TextRegexFunctions
{
    private const string Deviation =
        "Uses .NET regular expressions, not PCRE2: \\d and \\w also match non-ASCII digits and letters, and PCRE-only syntax such as possessive quantifiers is not available.";

    // Patterns are not compiled (AOT); the static Regex methods cache recently used patterns.
    private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(1);

    public static void Register(FunctionRegistry registry)
    {
        ArgumentKind[] values = [ArgumentKind.Value];
        registry.Add(new FunctionInfo("REGEXTEST", 2, 3, values, Test) { Status = FunctionStatus.KnownDeviation, Deviation = Deviation });
        registry.Add(new FunctionInfo("REGEXEXTRACT", 2, 4, values, Extract) { Status = FunctionStatus.KnownDeviation, Deviation = Deviation });
        registry.Add(new FunctionInfo("REGEXREPLACE", 3, 5, values, Replace) { Status = FunctionStatus.KnownDeviation, Deviation = Deviation });
    }

    private static CellValue Error(ErrorKind kind) => CellValue.Error(kind);

    // The case_sensitivity argument: 0 (default) matches case, 1 ignores it. Matching never
    // depends on the process culture.
    private static CellValue Options(FunctionCall call, int index, out RegexOptions options)
    {
        options = RegexOptions.CultureInvariant;
        var flag = call.Integer(index, 0);
        if (flag.IsError)
            return flag;
        if (flag.AsNumber() is not (0 or 1))
            return Error(ErrorKind.Value);
        if (flag.AsNumber() == 1)
            options |= RegexOptions.IgnoreCase;
        return flag;
    }

    private static Operand Run(Func<CellValue> body)
    {
        try
        {
            return body();
        }
        catch (ArgumentException)
        {
            // Includes RegexParseException: the pattern is not a valid expression.
            return Error(ErrorKind.Value);
        }
        catch (RegexMatchTimeoutException)
        {
            return Error(ErrorKind.Value);
        }
    }

    private static Operand Test(FunctionCall call)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        var pattern = call.Text(1);
        if (pattern.IsError)
            return pattern;
        var flag = Options(call, 2, out var options);
        if (flag.IsError)
            return flag;
        return Run(() => CellValue.Boolean(Regex.IsMatch(text.AsText(), pattern.AsText(), options, Timeout)));
    }

    // REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity]): 0 the first match, 1 all
    // matches in a row, 2 the capture groups of the first match in a row. No match is #N/A.
    private static Operand Extract(FunctionCall call)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        var pattern = call.Text(1);
        if (pattern.IsError)
            return pattern;
        var mode = call.Integer(2, 0);
        if (mode.IsError)
            return mode;
        if (mode.AsNumber() is not (0 or 1 or 2))
            return Error(ErrorKind.Value);
        var flag = Options(call, 3, out var options);
        if (flag.IsError)
            return flag;

        return Run(() =>
        {
            var input = text.AsText();
            switch (mode.AsNumber())
            {
                case 0:
                    var match = Regex.Match(input, pattern.AsText(), options, Timeout);
                    return match.Success ? TextFunctions.Result(match.Value) : Error(ErrorKind.NA);
                case 1:
                    var matches = Regex.Matches(input, pattern.AsText(), options, Timeout);
                    if (matches.Count == 0)
                        return Error(ErrorKind.NA);
                    var all = new CellValue[1, matches.Count];
                    for (var i = 0; i < matches.Count; i++)
                    {
                        call.Context.CancellationToken.ThrowIfCancellationRequested();
                        all[0, i] = CellValue.Text(matches[i].Value);
                    }

                    return all.Length == 1 ? all[0, 0] : CellValue.Array(all);
                default:
                    var first = Regex.Match(input, pattern.AsText(), options, Timeout);
                    if (!first.Success || first.Groups.Count < 2)
                        return Error(ErrorKind.NA);
                    var groups = new CellValue[1, first.Groups.Count - 1];
                    for (var i = 1; i < first.Groups.Count; i++)
                        groups[0, i - 1] = CellValue.Text(first.Groups[i].Success ? first.Groups[i].Value : "");
                    return groups.Length == 1 ? groups[0, 0] : CellValue.Array(groups);
            }
        });
    }

    // REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity]): occurrence 0
    // replaces every match, n the n-th, -n the n-th from the end. $1 and ${name} insert groups.
    private static Operand Replace(FunctionCall call)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        var pattern = call.Text(1);
        if (pattern.IsError)
            return pattern;
        var replacement = call.Text(2);
        if (replacement.IsError)
            return replacement;
        var occurrence = call.Integer(3, 0);
        if (occurrence.IsError)
            return occurrence;
        var flag = Options(call, 4, out var options);
        if (flag.IsError)
            return flag;

        return Run(() =>
        {
            var input = text.AsText();
            var n = occurrence.AsNumber();
            if (n == 0)
                return TextFunctions.Result(Regex.Replace(input, pattern.AsText(), replacement.AsText(), options, Timeout));

            var matches = Regex.Matches(input, pattern.AsText(), options, Timeout);
            var index = n > 0 ? n - 1 : matches.Count + n;
            if (index < 0 || index >= matches.Count)
                return text;
            var match = matches[(int)index];
            return TextFunctions.Result(string.Concat(input.AsSpan(0, match.Index), match.Result(replacement.AsText()), input.AsSpan(match.Index + match.Length)));
        });
    }
}
