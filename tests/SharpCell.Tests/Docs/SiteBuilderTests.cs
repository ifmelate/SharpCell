using SharpCell.Docs;
using SharpCell.Tests.Xlsx;

namespace SharpCell.Tests.Docs;

public class SiteBuilderTests
{
    [Fact]
    public void Links_to_markdown_pages_point_at_html_and_keep_their_anchors()
    {
        var html = SiteBuilder.RenderPage(
            "See [formulas](guides/formulas.md#spill), [API](api/index.md), [the site](https://example.org/x.md), "
            + "[data](compatibility.json), [mail](mailto:a@b.c) and [here](#top).",
            "index.md", "Overview", "");

        Assert.Contains("href=\"guides/formulas.html#spill\"", html);
        Assert.Contains("href=\"api/index.html\"", html);
        Assert.Contains("href=\"https://example.org/x.md\"", html);
        Assert.Contains("href=\"compatibility.json\"", html);
        Assert.Contains("href=\"mailto:a@b.c\"", html);
        Assert.Contains("href=\"#top\"", html);
    }

    [Fact]
    public void Links_to_repository_files_outside_docs_point_at_the_repository()
    {
        const string markdown = "See [the sample](../samples/SharpCell.Sample/README.md) and [its project](../../samples/SharpCell.Sample/SharpCell.Sample.csproj).";

        var top = SiteBuilder.RenderPage("See [the sample](../samples/SharpCell.Sample/README.md).", "getting-started.md", "T", "", "https://github.com/o/r");
        Assert.Contains("href=\"https://github.com/o/r/blob/main/samples/SharpCell.Sample/README.md\"", top);

        var nested = SiteBuilder.RenderPage(markdown, "guides/x.md", "T", "", "https://github.com/o/r/");
        Assert.Contains("href=\"https://github.com/o/r/blob/main/samples/SharpCell.Sample/SharpCell.Sample.csproj\"", nested);

        var unknown = SiteBuilder.RenderPage("See [the sample](../samples/SharpCell.Sample/README.md).", "getting-started.md", "T", "", "");
        Assert.DoesNotContain("samples/SharpCell.Sample", unknown);
        Assert.Contains("See the sample.", unknown[unknown.IndexOf("<main>")..unknown.IndexOf("<footer>")]);
    }

    [Fact]
    public void A_page_links_to_its_markdown_source_and_to_llms_txt_relative_to_its_folder()
    {
        var html = SiteBuilder.RenderPage("# Formulas\n\nText.", "guides/formulas.md", "Formulas", "");
        Assert.Contains("<title>Formulas · SharpCell</title>", html);
        Assert.Contains("<link rel=\"stylesheet\" href=\"../style.css\">", html);
        Assert.Contains("<link rel=\"alternate\" type=\"text/markdown\" href=\"formulas.md\">", html);
        Assert.Contains("href=\"../llms.txt\"", html);
    }

    [Fact]
    public void Tables_code_and_raw_anchors_render()
    {
        var html = SiteBuilder.RenderPage("<a id=\"x\"></a>\n### X\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n```csharp\nvar a = 1;\n```\n", "api/T.md", "T", "");
        Assert.Contains("<a id=\"x\"></a>", html);
        Assert.Contains("<table>", html);
        Assert.Contains("<code class=\"language-csharp\">", html);
    }

    [Fact]
    public void Navigation_marks_the_current_page_and_is_relative_to_it()
    {
        var config = SiteConfig.Parse("""
            { "title": "SharpCell", "summary": "s", "details": "", "baseUrl": "",
              "sections": [ { "title": "Docs", "pages": [
                { "path": "index.md", "title": "Overview", "description": "d" },
                { "path": "guides/formulas.md", "title": "Formulas", "description": "d" } ] } ],
              "optional": [] }
            """);

        var nav = SiteBuilder.Navigation(config, "guides/formulas.md");
        Assert.Contains("<a href=\"../index.html\">Overview</a>", nav);
        Assert.Contains("<a href=\"formulas.html\" aria-current=\"page\">Formulas</a>", nav);
        Assert.Contains("<a href=\"../api/index.html\">API reference</a>", nav);
        Assert.Contains("<a href=\"../compatibility.html\">Compatibility</a>", nav);
    }

    [Fact]
    public void The_built_site_has_html_and_markdown_for_every_page()
    {
        var output = Path.Combine(Path.GetTempPath(), "sharpcell-site-" + Guid.NewGuid().ToString("N"));
        try
        {
            SiteBuilder.Build(RepositoryLayout.At(Path.Combine(CorpusFiles.Root, "..", "..")), output);
            foreach (var page in new[] { "index", "getting-started", "guides/formulas", "agents", "compatibility", "api/index", "api/SharpCell.Workbook" })
            {
                Assert.True(File.Exists(Path.Combine(output, page + ".html")), page + ".html");
                Assert.True(File.Exists(Path.Combine(output, page + ".md")), page + ".md");
            }

            foreach (var file in new[] { "llms.txt", "llms-full.txt", "compatibility.json", "style.css", ".nojekyll" })
                Assert.True(File.Exists(Path.Combine(output, file)), file);
        }
        finally
        {
            if (Directory.Exists(output))
                Directory.Delete(output, recursive: true);
        }
    }
}
