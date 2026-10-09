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
