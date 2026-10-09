using SharpCell.Docs;

namespace SharpCell.Tests.Docs;

public class SnippetsTests
{
    private static Dictionary<string, string> Collect(string source) => Snippets.Collect([("Example.cs", source)]);

    [Fact]
    public void A_snippet_is_the_lines_between_its_markers_without_common_indentation()
    {
        var snippets = Collect("""
            class C
            {
                void M()
                {
                    // snippet: hello
                    var a = 1;
                    if (a > 0)
                        a++;
                    // end-snippet
                }
            }
            """);

        Assert.Equal("var a = 1;\nif (a > 0)\n    a++;", snippets["hello"]);
    }

    [Fact]
    public void A_snippet_name_used_twice_is_an_error()
    {
        var error = Assert.Throws<DocsException>(() => Collect("// snippet: a\n// end-snippet\n// snippet: a\n// end-snippet\n"));
        Assert.Contains("'a' is defined twice", error.Message);
    }

    [Fact]
    public void An_unclosed_snippet_is_an_error()
    {
        var error = Assert.Throws<DocsException>(() => Collect("// snippet: a\nvar x = 1;\n"));
        Assert.Contains("Example.cs:1", error.Message);
        Assert.Contains("'a' has no '// end-snippet'", error.Message);
    }

    [Fact]
    public void A_snippet_that_starts_inside_another_is_an_error()
    {
        var error = Assert.Throws<DocsException>(() => Collect("// snippet: a\n// snippet: b\n// end-snippet\n"));
        Assert.Contains("'a' is not closed", error.Message);
    }

    [Fact]
    public void Apply_replaces_the_body_of_a_named_block()
    {
        var markdown = "# Title\n\n```csharp snippet=hello\nold code\n```\n\nText.\n";
        var result = Snippets.Apply(markdown, new Dictionary<string, string> { ["hello"] = "var a = 1;\nvar b = 2;" }, "page.md");
        Assert.Equal("# Title\n\n```csharp snippet=hello\nvar a = 1;\nvar b = 2;\n```\n\nText.\n", result);
    }

    [Fact]
    public void Apply_leaves_other_code_blocks_alone()
    {
        var markdown = "```csharp\nnot a snippet\n```\n\n```bash\ndotnet run\n```\n";
        Assert.Equal(markdown, Snippets.Apply(markdown, new Dictionary<string, string>(), "page.md"));
    }

    [Fact]
    public void Apply_keeps_windows_line_endings()
    {
        var markdown = "Intro\r\n\r\n```csharp snippet=hello\r\nold\r\n```\r\n";
        var result = Snippets.Apply(markdown, new Dictionary<string, string> { ["hello"] = "a();\nb();" }, "page.md");
        Assert.Equal("Intro\r\n\r\n```csharp snippet=hello\r\na();\r\nb();\r\n```\r\n", result);
    }

    [Fact]
    public void Apply_rejects_an_unknown_snippet_with_the_file_and_line()
    {
        var error = Assert.Throws<DocsException>(() =>
            Snippets.Apply("Text\n\n```csharp snippet=missing\n```\n", new Dictionary<string, string>(), "docs/page.md"));
        Assert.Equal("docs/page.md:3: unknown snippet 'missing'.", error.Message);
    }

    [Fact]
    public void Apply_rejects_an_unclosed_block()
    {
        var error = Assert.Throws<DocsException>(() =>
            Snippets.Apply("```csharp snippet=hello\ncode\n", new Dictionary<string, string> { ["hello"] = "x" }, "page.md"));
        Assert.Contains("is not closed", error.Message);
    }
}
