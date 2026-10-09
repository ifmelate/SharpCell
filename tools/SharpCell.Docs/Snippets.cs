using System;
using System.Collections.Generic;
using System.Linq;

namespace SharpCell.Docs;

/// <summary>
/// Code examples live in compiled, tested .cs files between <c>// snippet: name</c> and
/// <c>// end-snippet</c>. A Markdown code block whose info string is <c>csharp snippet=name</c>
/// gets that code as its body, so the documentation shows only code that builds and passes.
/// </summary>
internal static class Snippets
{
    private const string Begin = "// snippet:";
    private const string End = "// end-snippet";
    private const string Fence = "```csharp snippet=";

    public static Dictionary<string, string> Collect(IEnumerable<(string Path, string Text)> sources)
    {
        var snippets = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (path, text) in sources)
        {
            var lines = Lines(text);
            for (var i = 0; i < lines.Length; i++)
            {
                var trimmed = lines[i].Trim();
                if (!trimmed.StartsWith(Begin, StringComparison.Ordinal))
                    continue;

                var name = trimmed[Begin.Length..].Trim();
                if (name.Length == 0)
                    throw new DocsException($"{path}:{i + 1}: a snippet needs a name.");

                var body = new List<string>();
                var j = i + 1;
                for (; j < lines.Length && lines[j].Trim() != End; j++)
                {
                    if (lines[j].Trim().StartsWith(Begin, StringComparison.Ordinal))
                        throw new DocsException($"{path}:{j + 1}: snippet '{name}' is not closed before the next one starts.");
                    body.Add(lines[j]);
                }

                if (j == lines.Length)
                    throw new DocsException($"{path}:{i + 1}: snippet '{name}' has no '{End}'.");
                if (!snippets.TryAdd(name, Dedent(body)))
                    throw new DocsException($"{path}:{i + 1}: snippet '{name}' is defined twice.");
                i = j;
            }
        }

        return snippets;
    }

    public static string Apply(string markdown, IReadOnlyDictionary<string, string> snippets, string path)
    {
        var newline = markdown.Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : "\n";
        var lines = Lines(markdown);
        var output = new List<string>(lines.Length);
        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            if (!line.StartsWith(Fence, StringComparison.Ordinal))
            {
                output.Add(line);
                continue;
            }

            var name = line[Fence.Length..].Trim();
            if (!snippets.TryGetValue(name, out var body))
                throw new DocsException($"{path}:{i + 1}: unknown snippet '{name}'.");

            var close = i + 1;
            while (close < lines.Length && lines[close].TrimEnd() != "```")
                close++;
            if (close == lines.Length)
                throw new DocsException($"{path}:{i + 1}: the code block for snippet '{name}' is not closed.");

            output.Add(line);
            if (body.Length > 0)
                output.AddRange(Lines(body));
            output.Add("```");
            i = close;
        }

        return string.Join(newline, output);
    }

    private static string[] Lines(string text) => text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');

    private static string Dedent(List<string> lines)
    {
        var indent = lines.Where(l => l.Trim().Length > 0).Select(l => l.Length - l.TrimStart().Length).DefaultIfEmpty(0).Min();
        return string.Join("\n", lines.Select(l => l.Trim().Length == 0 ? "" : l[indent..].TrimEnd()));
    }
}
