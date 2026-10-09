using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Markdig;
using Markdig.Renderers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace SharpCell.Docs;

/// <summary>
/// Renders docs/ as a static site: one HTML page per Markdown page, the Markdown next to it for
/// agents, llms.txt, llms-full.txt and compatibility.json. No JavaScript.
/// </summary>
internal static class SiteBuilder
{
    private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();

    public static void Build(RepositoryLayout layout, string output)
    {
        var config = SiteConfig.Load(layout.Full("docs/site.json"));
        var api = ApiReference.Generate(ApiReference.DocumentedAssemblies());
        var titles = config.Sections.SelectMany(s => s.Pages).ToDictionary(p => p.Path, p => p.Title, StringComparer.Ordinal);
        titles["compatibility.md"] = "Compatibility";
        titles["api/index.md"] = "API reference";
        foreach (var page in api.Pages)
            titles[page.Path["docs/".Length..]] = page.Type.Name;

        Directory.CreateDirectory(output);
        foreach (var (path, title) in titles.Where(t => t.Key.EndsWith(".md", StringComparison.Ordinal)))
        {
            var markdown = File.ReadAllText(layout.Full("docs/" + path));
            var target = Path.Combine(output, path.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            File.WriteAllText(target, markdown);
            File.WriteAllText(Path.ChangeExtension(target, ".html"), RenderPage(markdown, path, title, Navigation(config, path), config.RepositoryUrl ?? ""));
        }

        foreach (var file in new[] { "llms.txt", "llms-full.txt", "compatibility.json" })
            File.Copy(layout.Full("docs/" + file), Path.Combine(output, file), overwrite: true);
        File.Copy(layout.Full("tools/SharpCell.Docs/assets/style.css"), Path.Combine(output, "style.css"), overwrite: true);
        File.WriteAllText(Path.Combine(output, ".nojekyll"), "");
    }

    public static string Navigation(SiteConfig config, string pagePath)
    {
        var sb = new StringBuilder("<nav>\n");
        foreach (var section in config.Sections)
        {
            sb.Append("<p>").Append(WebUtility.HtmlEncode(section.Title)).Append("</p>\n<ul>\n");
            foreach (var page in section.Pages)
                Item(sb, page.Path, page.Title, pagePath);
            sb.Append("</ul>\n");
        }

        sb.Append("<p>Reference</p>\n<ul>\n");
        Item(sb, "api/index.md", "API reference", pagePath);
        Item(sb, "compatibility.md", "Compatibility", pagePath);
        sb.Append("</ul>\n</nav>");
        return sb.ToString();
    }

    /// <param name="repositoryUrl">
    /// Where links to files outside docs/ (the sample, the project file) point: the repository's web
    /// address. Empty until the repository is public; such links then keep only their text.
    /// </param>
    public static string RenderPage(string markdown, string pagePath, string title, string navigation, string repositoryUrl = "")
    {
        var document = Markdown.Parse(markdown, Pipeline);
        foreach (var link in document.Descendants<LinkInline>().ToList())
        {
            if (link.Url is not { } url || IsExternal(url))
                continue;
            if (OutsideDocs(pagePath, url) is { } repositoryPath)
            {
                if (repositoryUrl.Length > 0)
                    link.Url = repositoryUrl.TrimEnd('/') + "/blob/main/" + repositoryPath;
                else
                    link.ReplaceBy(new LiteralInline(Text(link)), copyChildren: false);
            }
            else if (IsLocalMarkdown(url))
            {
                link.Url = ToHtml(url);
            }
        }

        using var writer = new StringWriter();
        var renderer = new HtmlRenderer(writer);
        Pipeline.Setup(renderer);
        renderer.Render(document);
        writer.Flush();

        var root = Root(pagePath);
        var source = Path.GetFileName(pagePath);
        return $"""
            <!doctype html>
            <html lang="en">
            <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <title>{WebUtility.HtmlEncode(title)} · SharpCell</title>
            <link rel="stylesheet" href="{root}style.css">
            <link rel="alternate" type="text/markdown" href="{source}">
            </head>
            <body>
            <header><a class="home" href="{root}index.html">SharpCell</a></header>
            <div class="layout">
            {navigation}
            <main>
            {writer}
            <footer><a href="{source}">Markdown source</a> · <a href="{root}llms.txt">llms.txt</a> · <a href="{root}llms-full.txt">llms-full.txt</a></footer>
            </main>
            </div>
            </body>
            </html>

            """;
    }

    private static void Item(StringBuilder sb, string path, string title, string current)
    {
        var href = Relative(current, ToHtml(path));
        sb.Append("<li><a href=\"").Append(href).Append('"');
        if (path == current)
            sb.Append(" aria-current=\"page\"");
        sb.Append('>').Append(WebUtility.HtmlEncode(title)).Append("</a></li>\n");
    }

    private static bool IsExternal(string url) =>
        url.StartsWith('#') || url.Contains("://", StringComparison.Ordinal) || url.StartsWith("mailto:", StringComparison.Ordinal);

    /// <summary>The repository path a relative link leads to when it leaves docs/, otherwise null.</summary>
    private static string? OutsideDocs(string pagePath, string url)
    {
        var segments = new List<string>(pagePath.Split('/')[..^1]);
        foreach (var segment in url.Split('#')[0].Split('/'))
        {
            if (segment == "..")
            {
                if (segments.Count == 0)
                    segments.Insert(0, "..");
                else if (segments[^1] == "..")
                    segments.Add("..");
                else
                    segments.RemoveAt(segments.Count - 1);
            }
            else if (segment is not ("." or ""))
            {
                segments.Add(segment);
            }
        }

        // Exactly one ".." left: the link climbed out of docs/ into the repository root.
        return segments.Count > 1 && segments[0] == ".." && segments[1] != ".." ? string.Join('/', segments.Skip(1)) : null;
    }

    private static string Text(Inline inline) => inline switch
    {
        LiteralInline literal => literal.Content.ToString(),
        CodeInline code => code.Content,
        ContainerInline container => string.Concat(container.Select(Text)),
        _ => "",
    };

    private static bool IsLocalMarkdown(string url)
    {
        if (IsExternal(url))
            return false;
        var hash = url.IndexOf('#', StringComparison.Ordinal);
        var path = hash >= 0 ? url[..hash] : url;
        return path.EndsWith(".md", StringComparison.Ordinal);
    }

    private static string ToHtml(string url)
    {
        var hash = url.IndexOf('#', StringComparison.Ordinal);
        var path = hash >= 0 ? url[..hash] : url;
        return path[..^".md".Length] + ".html" + (hash >= 0 ? url[hash..] : "");
    }

    private static string Root(string pagePath) => string.Concat(Enumerable.Repeat("../", pagePath.Count(c => c == '/')));

    // A path relative to the site root, rewritten relative to the current page's folder.
    private static string Relative(string current, string target)
    {
        var currentFolder = current.Contains('/', StringComparison.Ordinal) ? current[..current.LastIndexOf('/')] + "/" : "";
        return currentFolder.Length > 0 && target.StartsWith(currentFolder, StringComparison.Ordinal)
            ? target[currentFolder.Length..]
            : Root(current) + target;
    }
}
