using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SharpCell.Docs;

/// <summary>
/// llms.txt (https://llmstxt.org): a Markdown index an agent reads first, and llms-full.txt, every
/// page in one file for agents that take a whole document.
/// </summary>
internal static partial class LlmsText
{
    public static string Index(SiteConfig config, IReadOnlyList<ApiPage> api)
    {
        var sb = new StringBuilder();
        Header(sb, config);
        foreach (var section in config.Sections)
        {
            sb.Append("\n## ").Append(section.Title).Append("\n\n");
            foreach (var page in section.Pages)
                sb.Append("- [").Append(page.Title).Append("](").Append(config.Url(page.Path)).Append("): ").Append(page.Description).Append('\n');
        }

        sb.Append("\n## API reference\n\n");
        foreach (var page in api)
            sb.Append("- [").Append(page.Type.Name).Append("](").Append(config.Url(DocsRelative(page.Path))).Append("): ").Append(MarkdownLink().Replace(page.Summary, "$1")).Append('\n');

        if (config.Optional.Count > 0)
        {
            sb.Append("\n## Optional\n\n");
            foreach (var page in config.Optional)
                sb.Append("- [").Append(page.Title).Append("](").Append(config.Url(page.Path)).Append("): ").Append(page.Description).Append('\n');
        }

        return sb.ToString();
    }

    public static string Full(SiteConfig config, IReadOnlyList<ApiPage> api, Func<string, string> read)
    {
        var sb = new StringBuilder();
        Header(sb, config);
        var paths = config.Sections.SelectMany(s => s.Pages).Select(p => p.Path)
            .Where(p => p.EndsWith(".md", StringComparison.Ordinal))
            .Concat(api.Select(p => DocsRelative(p.Path)));
        foreach (var path in paths)
        {
            var text = read(path).Replace("\r\n", "\n", StringComparison.Ordinal);
            if (text.StartsWith("<!--", StringComparison.Ordinal))
                text = text[(text.IndexOf("-->", StringComparison.Ordinal) + 3)..];
            sb.Append("\n---\n\n<!-- page: ").Append(path).Append(" -->\n\n").Append(text.Trim()).Append('\n');
        }

        return sb.ToString();
    }

    private static void Header(StringBuilder sb, SiteConfig config)
    {
        sb.Append("# ").Append(config.Title).Append("\n\n> ").Append(config.Summary).Append('\n');
        if (config.Details.Length > 0)
            sb.Append('\n').Append(config.Details).Append('\n');
    }

    // API summaries link relative to docs/api/; llms.txt sits in docs/, so keep the link text only.
    [GeneratedRegex(@"\[([^\]]*)\]\([^)]*\)")]
    private static partial Regex MarkdownLink();

    private static string DocsRelative(string repositoryPath) =>
        repositoryPath.StartsWith("docs/", StringComparison.Ordinal) ? repositoryPath["docs/".Length..] : repositoryPath;
}
