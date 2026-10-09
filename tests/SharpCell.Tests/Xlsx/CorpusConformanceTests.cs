using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using SharpCell.Conformance;

namespace SharpCell.Tests.Xlsx;

/// <summary>
/// A ratchet on the Excel corpus: no file may match Excel in fewer cells than the committed
/// compatibility report says, and files that match fully must stay so. Improvements pass; run
/// <c>dotnet run --project tools/SharpCell.Conformance</c> to record them in the report.
/// </summary>
public class CorpusConformanceTests
{
    private static readonly Dictionary<string, (int Cells, int Passed)> Baseline = ReadBaseline();

    public static TheoryData<string> Files() => [.. CorpusFiles.All];

    [Fact]
    public void Report_covers_every_corpus_file()
    {
        var corpus = new HashSet<string>(CorpusFiles.All, StringComparer.Ordinal);
        corpus.SymmetricExceptWith(Baseline.Keys);
        Assert.True(corpus.Count == 0, "Corpus files and docs/compatibility.json disagree; regenerate the report: " + string.Join(", ", corpus));
    }

    [Theory]
    [MemberData(nameof(Files))]
    public void File_matches_Excel_at_least_as_well_as_recorded(string file)
    {
        var result = CorpusRunner.Run(CorpusFiles.PathOf(file), file, TimeSpan.FromMinutes(2));
        Assert.Null(result.Error);
        if (!Baseline.TryGetValue(file, out var recorded))
            return;

        var failures = new List<string>();
        foreach (var cell in result.Cells)
        {
            if (!cell.Passed && failures.Count < 20)
                failures.Add($"{cell.Sheet}!{cell.Address} {cell.Formula ?? "(spilled)"}: Excel {cell.Expected}, SharpCell {cell.Actual}");
        }

        Assert.True(result.Passed >= recorded.Passed,
            $"{file}: {result.Passed} cells match Excel, the report records {recorded.Passed}. Mismatches:\n" + string.Join("\n", failures));
        if (recorded.Passed == recorded.Cells)
            Assert.True(result.Passed == result.Cells.Count, $"{file} matched Excel fully and no longer does:\n" + string.Join("\n", failures));
    }

    private static Dictionary<string, (int, int)> ReadBaseline()
    {
        var path = Path.Combine(CorpusFiles.Root, "..", "..", "docs", "compatibility.json");
        var result = new Dictionary<string, (int, int)>(StringComparer.Ordinal);
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        foreach (var file in document.RootElement.GetProperty("files").EnumerateArray())
            result[file.GetProperty("file").GetString()!] = (file.GetProperty("cells").GetInt32(), file.GetProperty("passed").GetInt32());
        return result;
    }
}
