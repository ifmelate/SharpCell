using SharpCell.Docs;
using SharpCell.Tests.Xlsx;

namespace SharpCell.Tests.Docs;

public class DocsUpToDateTests
{
    private static RepositoryLayout Repository() => RepositoryLayout.At(Path.Combine(CorpusFiles.Root, "..", ".."));

    [Fact]
    public void The_committed_documentation_matches_what_the_tool_generates()
    {
        var layout = Repository();
        var differences = Generator.Differences(layout, Generator.Run(layout));
        Assert.True(differences.Count == 0,
            "Run 'dotnet run --project tools/SharpCell.Docs -f net10.0 -- generate':\n" + string.Join("\n", differences));
    }

    [Fact]
    public void A_checkout_with_windows_line_endings_is_up_to_date()
    {
        var root = Path.Combine(Path.GetTempPath(), "sharpcell-docs-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(Path.Combine(root, "docs"));
            File.WriteAllText(Path.Combine(root, "docs", "llms.txt"), "# SharpCell\r\n\r\n> Summary.\r\n");
            var layout = RepositoryLayout.At(root);

            Assert.Empty(Generator.Differences(layout, [new GeneratedFile("docs/llms.txt", "# SharpCell\n\n> Summary.\n")]));
            Assert.Equal(["docs/llms.txt: out of date"], Generator.Differences(layout, [new GeneratedFile("docs/llms.txt", "# SharpCell\n\n> Other.\n")]));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void A_hand_edit_or_a_stray_file_in_the_api_folder_is_reported()
    {
        var layout = Repository();
        var files = Generator.Run(layout).ToList();
        var edited = files.Select(f => f.Path == "docs/api/SharpCell.Workbook.md" ? f with { Content = f.Content + "edit\n" } : f).ToList();
        Assert.Contains("docs/api/SharpCell.Workbook.md: out of date", Generator.Differences(layout, edited));

        var withoutCell = files.Where(f => f.Path != "docs/api/SharpCell.Cell.md").ToList();
        Assert.Contains("docs/api/SharpCell.Cell.md: not generated any more", Generator.Differences(layout, withoutCell));
    }
}
