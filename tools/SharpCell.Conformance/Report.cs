using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using SharpCell.Functions;

namespace SharpCell.Conformance;

internal sealed record FunctionStats(string Name, string Status, string? Deviation, int Cases, int Passed);

/// <summary>
/// The compatibility table: per corpus file and per function, how many compared cells match
/// Excel. Output is deterministic (no timestamps, ordinal sorting) so it can be committed and diffed.
/// </summary>
internal sealed class Report
{
    // Parsed as calls but handled by the evaluator itself rather than the function registry.
    private static readonly HashSet<string> SpecialForms = new(StringComparer.Ordinal) { "LET", "LAMBDA" };

    public Report(IReadOnlyList<FileResult> files, CorpusOverrides? overrides = null)
    {
        Overrides = overrides ?? CorpusOverrides.Empty;
        Files = [.. files.OrderBy(f => f.File, StringComparer.Ordinal)];
        var cases = new Dictionary<string, (int Cases, int Passed)>(StringComparer.Ordinal);
        foreach (var file in Files)
        {
            foreach (var cell in file.Cells)
            {
                foreach (var function in cell.Functions)
                {
                    var (total, passed) = cases.GetValueOrDefault(function);
                    cases[function] = (total + 1, passed + (cell.Passed ? 1 : 0));
                }
            }
        }

        var names = new SortedSet<string>(cases.Keys, StringComparer.Ordinal);
        foreach (var function in FunctionRegistry.Default.All)
            names.Add(function.Name);

        Functions = [.. names.Select(name =>
        {
            var (total, passed) = cases.GetValueOrDefault(name);
            var (status, deviation) = StatusOf(name);
            return new FunctionStats(name, status, deviation, total, passed);
        })];
    }

    public IReadOnlyList<FileResult> Files { get; }

    public CorpusOverrides Overrides { get; }

    /// <summary>Matching cells that match only within a tolerance stated in the overrides file.</summary>
    public int Widened => Files.Sum(f => f.Cells.Count(c => c.Widened));

    public IReadOnlyList<FunctionStats> Functions { get; }

    public int Cells => Files.Sum(f => f.Cells.Count);

    public int Passed => Files.Sum(f => f.Passed);

    private static (string Status, string? Deviation) StatusOf(string name)
    {
        if (SpecialForms.Contains(name))
            return ("implemented", null);
        if (!FunctionRegistry.Default.TryGet(name, out var info))
            return ("not implemented", null);
        return info!.Status switch
        {
            FunctionStatus.Implemented => ("implemented", null),
            FunctionStatus.KnownDeviation => ("known deviation", info.Deviation),
            _ => ("not implemented", null),
        };
    }

    public string ToJson()
    {
        using var stream = new MemoryStream();
        using (var json = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true }))
        {
            json.WriteStartObject();
            json.WriteNumber("cells", Cells);
            json.WriteNumber("passed", Passed);
            json.WriteNumber("widened", Widened);
            json.WriteStartArray("files");
            foreach (var file in Files)
            {
                json.WriteStartObject();
                json.WriteString("file", file.File);
                json.WriteNumber("cells", file.Cells.Count);
                json.WriteNumber("passed", file.Passed);
                json.WriteNumber("skipped", file.Skipped);
                json.WriteNumber("widened", file.Cells.Count(c => c.Widened));
                if (file.Error is not null)
                    json.WriteString("error", file.Error);

                // Exactly which cells are compared and which of them differ, so the ratchet sees a
                // cell that breaks even when another one in the same file gets fixed.
                WriteCells(json, "comparedCells", file.Cells);
                WriteCells(json, "failingCells", file.Cells.Where(c => !c.Passed));
                json.WriteEndObject();
            }

            json.WriteEndArray();
            json.WriteStartArray("functions");
            foreach (var function in Functions)
            {
                json.WriteStartObject();
                json.WriteString("name", function.Name);
                json.WriteString("status", function.Status);
                if (function.Deviation is not null)
                    json.WriteString("deviation", function.Deviation);
                json.WriteNumber("cases", function.Cases);
                json.WriteNumber("passed", function.Passed);
                json.WriteEndObject();
            }

            json.WriteEndArray();
            json.WriteEndObject();
        }

        return Encoding.UTF8.GetString(stream.ToArray()).ReplaceLineEndings("\n") + "\n";
    }

    private static void WriteCells(Utf8JsonWriter json, string name, IEnumerable<CellOutcome> cells)
    {
        json.WriteStartObject(name);
        foreach (var (sheet, list) in CellSet.Encode(cells.Select(c => (c.Sheet, c.Cell))))
            json.WriteString(sheet, list);
        json.WriteEndObject();
    }

    public string ToMarkdown()
    {
        var md = new StringBuilder();
        md.Append("# Compatibility\n\n");
        md.Append("Generated by `tools/SharpCell.Conformance` from the reference workbooks in `tests/corpus` ");
        md.Append("(workbooks calculated by Microsoft Excel). Each formula result, and each cell an array formula ");
        md.Append("fills, is recalculated by SharpCell and compared with the value Excel saved: numbers within a ");
        md.Append("relative tolerance of 1e-9 (or the file's own, see `tests/corpus/ironcalc/SOURCE.md`), errors by ");
        md.Append("kind, text exactly. Random functions are compared by result type only. Exceptions are stated, with ");
        md.Append("their reasons, at the end.\n\n");
        md.Append(CultureInfo.InvariantCulture, $"**{Passed} of {Cells} cells match Excel ({Headline(Passed, Cells)}).**");
        if (Widened > 0)
            md.Append(CultureInfo.InvariantCulture, $" {Widened} of them match within a tolerance stated below, where Excel's own result is only as accurate as its solver.");
        md.Append("\n\n");

        var implemented = Functions.Count(f => f.Status == "implemented");
        md.Append("## Functions\n\n");
        md.Append(CultureInfo.InvariantCulture, $"{implemented} implemented, {Functions.Count(f => f.Status == "known deviation")} with a known deviation, ");
        md.Append(CultureInfo.InvariantCulture, $"{Functions.Count(f => f.Status == "not implemented")} used in the corpus but not implemented. ");
        md.Append("A case is a compared cell whose formula uses the function, so a failing case can be caused by another function in the same formula.\n\n");
        md.Append("| Function | Status | Cases | Match | % |\n|---|---|---:|---:|---:|\n");
        foreach (var function in Functions)
        {
            var status = function.Deviation is null ? function.Status : $"{function.Status}: {function.Deviation}";
            md.Append(CultureInfo.InvariantCulture,
                $"| {function.Name} | {Escape(status)} | {function.Cases} | {function.Passed} | {Percent(function.Passed, function.Cases)} |\n");
        }

        md.Append("\n## Files\n\n| File | Cells | Match | % | Note |\n|---|---:|---:|---:|---|\n");
        foreach (var file in Files)
        {
            md.Append(CultureInfo.InvariantCulture,
                $"| {Escape(file.File)} | {file.Cells.Count} | {file.Passed} | {Percent(file.Passed, file.Cells.Count)} | {Escape(file.Error ?? "")} |\n");
        }

        if (Overrides.Entries.Count > 0)
        {
            md.Append("\n## Stated exceptions\n\n");
            md.Append(CultureInfo.InvariantCulture, $"From `tests/corpus/{CorpusOverrides.FileName}`: the locale Excel calculated a file in, which a file does not record, ");
            md.Append("and wider tolerances for cells where Excel's result is only as accurate as its solver.\n\n");
            md.Append("| File | Cells | Rule | Reason |\n|---|---|---|---|\n");
            foreach (var entry in Overrides.Entries)
            {
                var rule = entry.Culture is { } culture
                    ? $"culture {culture}"
                    : string.Join(", ", new[]
                    {
                        entry.Absolute is { } a ? "absolute " + a.ToString("g", CultureInfo.InvariantCulture) : null,
                        entry.Relative is { } r ? "relative " + r.ToString("g", CultureInfo.InvariantCulture) : null,
                    }.Where(p => p is not null));
                md.Append(CultureInfo.InvariantCulture,
                    $"| {Escape(entry.File)} | {Escape(entry.Cells.Count == 0 ? "all" : string.Join(", ", entry.Cells))} | {rule} | {Escape(entry.Reason)} |\n");
            }
        }

        return md.ToString();
    }

    // Rounded down to a thousandth, so a corpus with failing cells never reads 100%.
    private static string Headline(int part, int whole) =>
        whole == 0 ? "–" : (Math.Floor(100_000.0 * part / whole) / 1000).ToString("0.000", CultureInfo.InvariantCulture) + "%";

    private static string Percent(int part, int whole) =>
        whole == 0 ? "–" : (100.0 * part / whole).ToString("0.0", CultureInfo.InvariantCulture) + "%";

    private static string Escape(string text) => text.Replace("|", "\\|", StringComparison.Ordinal).Replace('\n', ' ');
}
