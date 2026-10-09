using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.Json;
using SharpCell.Evaluation;

namespace SharpCell.Conformance;

/// <summary>
/// A stated exception to the default comparison, with its reason: the culture Excel calculated a
/// file in, or a wider tolerance for cells where Excel's own result is only as accurate as its
/// solver. Cells are <c>Sheet!A1</c> or <c>Sheet!A1:B2</c>.
/// </summary>
internal sealed record CorpusOverride(string File, string? Culture, IReadOnlyList<string> Cells, double? Absolute, double? Relative, string Reason);

/// <summary>The rules in <c>tests/corpus/overrides.json</c>; every rule has a reason and an effect.</summary>
internal sealed class CorpusOverrides
{
    public const string FileName = "overrides.json";

    private readonly List<(CorpusOverride Rule, string Sheet, Area Area)> _cells = [];

    private CorpusOverrides(IReadOnlyList<CorpusOverride> entries)
    {
        Entries = entries;
        foreach (var entry in entries)
        {
            foreach (var cells in entry.Cells)
                _cells.Add((entry, SheetOf(cells), AreaOf(cells)));
        }
    }

    public static CorpusOverrides Empty { get; } = new([]);

    public IReadOnlyList<CorpusOverride> Entries { get; }

    /// <summary>The rules of a corpus folder, or none when it has no overrides file.</summary>
    public static CorpusOverrides Load(string corpusRoot)
    {
        var path = Path.Combine(corpusRoot, FileName);
        return File.Exists(path) ? Parse(File.ReadAllText(path)) : Empty;
    }

    /// <exception cref="FormatException">A rule has no file, no reason or no effect, or names cells badly.</exception>
    public static CorpusOverrides Parse(string json)
    {
        using var document = JsonDocument.Parse(json);
        var entries = new List<CorpusOverride>();
        foreach (var element in document.RootElement.EnumerateArray())
        {
            var file = Text(element, "file") ?? throw new FormatException("A rule has no file.");
            var reason = Text(element, "reason") ?? throw new FormatException($"A rule for {file} has no reason.");
            var culture = Text(element, "culture");
            var cells = new List<string>();
            if (element.TryGetProperty("cells", out var list))
            {
                foreach (var item in list.EnumerateArray())
                    cells.Add(item.GetString() ?? throw new FormatException($"A cell of a rule for {file} is not text."));
            }

            double? absolute = element.TryGetProperty("absolute", out var a) ? a.GetDouble() : null;
            double? relative = element.TryGetProperty("relative", out var r) ? r.GetDouble() : null;
            if (culture is null && cells.Count == 0)
                throw new FormatException($"A rule for {file} changes nothing: give a culture or cells.");
            if (cells.Count > 0 && absolute is null && relative is null)
                throw new FormatException($"A rule for {file} names cells but no tolerance.");

            var entry = new CorpusOverride(file, culture, cells, absolute, relative, reason);
            foreach (var cell in cells)
            {
                SheetOf(cell);
                AreaOf(cell);
            }

            entries.Add(entry);
        }

        return new CorpusOverrides(entries);
    }

    public CultureInfo? CultureOf(string file)
    {
        foreach (var entry in Entries)
        {
            if (entry.File == file && entry.Culture is { } culture)
                return CultureInfo.GetCultureInfo(culture);
        }

        return null;
    }

    public CorpusOverride? For(string file, string sheet, CellAddress cell)
    {
        foreach (var (rule, ruleSheet, area) in _cells)
        {
            if (rule.File == file && ruleSheet == sheet && area.Contains(cell.Row, cell.Column))
                return rule;
        }

        return null;
    }

    // The sheet is everything before the last '!', so sheet names may hold spaces and dots.
    private static string SheetOf(string cells)
    {
        var bang = cells.LastIndexOf('!');
        return bang > 0 ? cells[..bang] : throw new FormatException($"'{cells}' names no sheet; write Sheet!A1.");
    }

    private static Area AreaOf(string cells)
    {
        var range = cells[(cells.LastIndexOf('!') + 1)..];
        var colon = range.IndexOf(':');
        var first = colon < 0 ? range : range[..colon];
        var last = colon < 0 ? range : range[(colon + 1)..];
        if (!CellAddress.TryParse(first, out var a) || !CellAddress.TryParse(last, out var b))
            throw new FormatException($"'{cells}' is not a cell or range such as Sheet!A1:B2.");
        return Area.Bounding(Area.Cell(a.Row, a.Column), Area.Cell(b.Row, b.Column));
    }

    private static string? Text(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;
}
