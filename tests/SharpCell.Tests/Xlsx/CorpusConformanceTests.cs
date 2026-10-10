using System;
using System.Collections.Generic;
using System.Linq;
using SharpCell.Conformance;

namespace SharpCell.Tests.Xlsx;

/// <summary>
/// A ratchet on the Excel corpus: every cell the committed compatibility report records as matching
/// Excel must still match. Improvements pass; run <c>dotnet run --project tools/SharpCell.Conformance</c>
/// to record them in the report.
/// </summary>
public class CorpusConformanceTests
{
    public static TheoryData<string> Files() => [.. CorpusFiles.All];

    [Fact]
    public void Report_covers_every_corpus_file()
    {
        var corpus = new HashSet<string>(CorpusFiles.All, StringComparer.Ordinal);
        corpus.SymmetricExceptWith(CorpusBaseline.Files);
        Assert.True(corpus.Count == 0, "Corpus files and docs/compatibility.json disagree; regenerate the report: " + string.Join(", ", corpus));
    }

    [Theory]
    [MemberData(nameof(Files))]
    public void Cells_that_matched_Excel_still_do(string file)
    {
        var result = CorpusRunner.Run(CorpusFiles.PathOf(file), file, TimeSpan.FromMinutes(2), CorpusOverrides.Load(CorpusFiles.Root));
        Assert.Null(result.Error);
        var broken = CorpusBaseline.Broken(file, result);
        Assert.True(broken.Count == 0, $"{file}: cells that matched Excel no longer do:\n" + string.Join("\n", broken.Take(30)));
    }
}
