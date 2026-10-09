using SharpCell.Docs;

namespace SharpCell.Tests.Docs;

public class LlmsTextTests
{
    private static readonly SiteConfig Config = SiteConfig.Parse("""
        {
          "title": "SharpCell",
          "summary": "An Excel formula engine for .NET.",
          "details": "Packages: SharpCell, SharpCell.Xlsx.",
          "baseUrl": "https://example.org/sharpcell/",
          "sections": [
            { "title": "Docs", "pages": [
              { "path": "index.md", "title": "Overview", "description": "What it is." },
              { "path": "guides/formulas.md", "title": "Formulas", "description": "Syntax and values." }
            ] }
          ],
          "optional": [ { "path": "compatibility.json", "title": "Compatibility data", "description": "Every function's status." } ]
        }
        """);

    private static readonly ApiPage[] Api = [new(typeof(Workbook), "docs/api/SharpCell.Workbook.md", "A workbook.")];

    [Fact]
    public void The_index_follows_the_llms_txt_format_with_absolute_links()
    {
        Assert.Equal("""
            # SharpCell

            > An Excel formula engine for .NET.

            Packages: SharpCell, SharpCell.Xlsx.

            ## Docs

            - [Overview](https://example.org/sharpcell/index.md): What it is.
            - [Formulas](https://example.org/sharpcell/guides/formulas.md): Syntax and values.

            ## API reference

            - [Workbook](https://example.org/sharpcell/api/SharpCell.Workbook.md): A workbook.

            ## Optional

            - [Compatibility data](https://example.org/sharpcell/compatibility.json): Every function's status.

            """.ReplaceLineEndings("\n"), LlmsText.Index(Config, Api));
    }

    [Fact]
    public void Without_a_base_url_links_are_relative()
    {
        var config = Config with { BaseUrl = "" };
        Assert.Contains("- [Overview](index.md): What it is.", LlmsText.Index(config, Api));
    }

    [Fact]
    public void Links_inside_api_summaries_become_plain_text_because_they_are_relative_to_the_api_folder()
    {
        ApiPage[] api = [new(typeof(CellValueKind), "docs/api/SharpCell.CellValueKind.md", "The kind of value a [CellValue](SharpCell.CellValue.md) holds.")];
        Assert.Contains("- [CellValueKind](https://example.org/sharpcell/api/SharpCell.CellValueKind.md): The kind of value a CellValue holds.\n",
            LlmsText.Index(Config, api));
    }

    [Fact]
    public void The_full_text_joins_the_pages_in_order_with_their_paths()
    {
        var pages = new Dictionary<string, string>
        {
            ["index.md"] = "# Overview\n\nHello.\n",
            ["guides/formulas.md"] = "# Formulas\n\nSyntax.\n",
            ["api/SharpCell.Workbook.md"] = "<!-- Generated -->\n\n# Workbook\n\nA workbook.\n",
        };

        var full = LlmsText.Full(Config, Api, path => pages[path]);

        Assert.StartsWith("# SharpCell\n\n> An Excel formula engine for .NET.\n", full);
        var overview = full.IndexOf("<!-- page: index.md -->", StringComparison.Ordinal);
        var formulas = full.IndexOf("<!-- page: guides/formulas.md -->", StringComparison.Ordinal);
        var workbook = full.IndexOf("<!-- page: api/SharpCell.Workbook.md -->", StringComparison.Ordinal);
        Assert.True(overview > 0 && overview < formulas && formulas < workbook);
        Assert.DoesNotContain("<!-- Generated -->", full);
    }

    [Fact]
    public void A_page_listed_in_the_config_must_exist()
    {
        var error = Assert.Throws<DocsException>(() => SiteConfig.Validate(Config with { Sections = [new("Docs", [new("nope.md", "Nope", "x")])] }, _ => false));
        Assert.Contains("nope.md", error.Message);
    }
}
