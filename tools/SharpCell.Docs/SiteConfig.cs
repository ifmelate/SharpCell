using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace SharpCell.Docs;

internal sealed record SitePage(string Path, string Title, string Description);

internal sealed record SiteSection(string Title, IReadOnlyList<SitePage> Pages);

/// <summary>docs/site.json: the site title, the order of the pages and their one-line descriptions.</summary>
internal sealed record SiteConfig(string Title, string Summary, string Details, string BaseUrl,
    IReadOnlyList<SiteSection> Sections, IReadOnlyList<SitePage> Optional)
{
    private static readonly JsonSerializerOptions Options = new() { PropertyNameCaseInsensitive = true };

    // Files the tool writes itself; they may be listed before they exist.
    private static readonly HashSet<string> Generated = new(StringComparer.Ordinal) { "llms.txt", "llms-full.txt" };

    public static SiteConfig Parse(string json) =>
        JsonSerializer.Deserialize<SiteConfig>(json, Options) ?? throw new DocsException("docs/site.json is empty.");

    public static SiteConfig Load(string path)
    {
        var config = Parse(File.ReadAllText(path));
        var docs = System.IO.Path.GetDirectoryName(path)!;
        Validate(config, page => File.Exists(System.IO.Path.Combine(docs, page)));
        return config;
    }

    public static void Validate(SiteConfig config, Func<string, bool> exists)
    {
        var missing = config.Sections.SelectMany(s => s.Pages).Concat(config.Optional)
            .Select(p => p.Path)
            .Where(p => !Generated.Contains(p) && !exists(p))
            .ToList();
        if (missing.Count > 0)
            throw new DocsException("docs/site.json lists pages that do not exist: " + string.Join(", ", missing));
    }

    public string Url(string path) => BaseUrl.Length == 0 ? path : BaseUrl.TrimEnd('/') + "/" + path;
}
