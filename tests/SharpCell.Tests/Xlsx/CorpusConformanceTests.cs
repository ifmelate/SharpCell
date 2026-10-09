using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using SharpCell.Conformance;

namespace SharpCell.Tests.Xlsx;

/// <summary>
/// A ratchet on the Excel corpus: every cell the committed compatibility report records as matching
/// Excel must still match. Improvements pass; run <c>dotnet run --project tools/SharpCell.Conformance</c>
/// to record them in the report.
/// </summary>
public class CorpusConformanceTests
{
    // The cells that matched Excel when the report was written: compared minus failing.
    private sealed record Recorded(int Cells, HashSet<(string, CellAddress)> Passing);

    private static readonly Dictionary<string, Recorded> Baseline = ReadBaseline();

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
    public void Cells_that_matched_Excel_still_do(string file)
    {
        var result = CorpusRunner.Run(CorpusFiles.PathOf(file), file, TimeSpan.FromMinutes(2));
        Assert.Null(result.Error);
        if (!Baseline.TryGetValue(file, out var recorded))
            return;

        var outcomes = result.Cells.ToDictionary(c => (c.Sheet, c.Cell));
        var broken = new List<string>();
        foreach (var cell in recorded.Passing)
        {
            if (!outcomes.TryGetValue(cell, out var outcome) || !outcome.Passed)
                broken.Add(Describe(cell, outcome));
        }

        Assert.True(broken.Count == 0, $"{file}: cells that matched Excel no longer do:\n" + string.Join("\n", broken.Take(30)));
    }

    private static string Describe((string Sheet, CellAddress Cell) cell, CellOutcome? outcome) => outcome is null
        ? $"{cell.Sheet}!{cell.Cell}: no longer compared"
        : $"{cell.Sheet}!{cell.Cell} {outcome.Formula ?? "(spilled)"}: Excel {outcome.Expected}, SharpCell {outcome.Actual}";

    private static HashSet<(string, CellAddress)> Decode(JsonElement sheets) =>
        CellSet.Decode(sheets.EnumerateObject().Select(p => KeyValuePair.Create(p.Name, p.Value.GetString()!)));

    private static Dictionary<string, Recorded> ReadBaseline()
    {
        var path = Path.Combine(CorpusFiles.Root, "..", "..", "docs", "compatibility.json");
        var result = new Dictionary<string, Recorded>(StringComparer.Ordinal);
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        foreach (var file in document.RootElement.GetProperty("files").EnumerateArray())
        {
            var compared = Decode(file.GetProperty("comparedCells"));
            compared.ExceptWith(Decode(file.GetProperty("failingCells")));
            result[file.GetProperty("file").GetString()!] = new Recorded(file.GetProperty("cells").GetInt32(), compared);
        }

        return result;
    }
}
