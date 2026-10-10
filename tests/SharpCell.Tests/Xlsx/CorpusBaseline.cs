using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using SharpCell.Conformance;

namespace SharpCell.Tests.Xlsx;

/// <summary>The cells the committed compatibility report records as matching Excel.</summary>
internal static class CorpusBaseline
{
    // The cells that matched Excel when the report was written: compared minus failing.
    private sealed record Recorded(int Cells, HashSet<(string, CellAddress)> Passing);

    private static readonly Dictionary<string, Recorded> Baseline = ReadBaseline();

    /// <summary>The corpus files the report covers.</summary>
    public static IEnumerable<string> Files => Baseline.Keys;

    /// <summary>Cells the report records as matching Excel that do not match in this result.</summary>
    public static IReadOnlyList<string> Broken(string file, FileResult result)
    {
        if (!Baseline.TryGetValue(file, out var recorded))
            return [];
        var outcomes = result.Cells.ToDictionary(c => (c.Sheet, c.Cell));
        var broken = new List<string>();
        foreach (var cell in recorded.Passing)
        {
            if (!outcomes.TryGetValue(cell, out var outcome) || !outcome.Passed)
                broken.Add(Describe(cell, outcome));
        }

        return broken;
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
