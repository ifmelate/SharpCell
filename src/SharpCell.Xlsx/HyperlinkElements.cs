using System.Collections.Generic;
using System.Xml;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Xlsx;

/// <summary>The attributes of one &lt;hyperlink&gt; element that decide what it means.</summary>
internal sealed record HyperlinkElement(string? Ref, string? RelationshipId, string? Location, string? Tooltip)
{
    /// <summary>The element the reader is on.</summary>
    public static HyperlinkElement Read(XmlReader reader) =>
        new(reader.GetAttribute("ref"), WorkbookReader.RelationshipId(reader), reader.GetAttribute("location"), reader.GetAttribute("tooltip"));
}

/// <summary>A &lt;hyperlink&gt; element the model holds: the cells it covers and the link they get.</summary>
internal readonly record struct ModelledHyperlink(Area Area, Hyperlink Link);

/// <summary>
/// Which &lt;hyperlink&gt; elements of a sheet the model holds. The reader applies those; the writer
/// keeps the others as the file has them, so both use this one decision.
/// </summary>
internal static class HyperlinkElements
{
    // A link over more cells than this is left out of the model.
    internal const long MaxCellsPerElement = 10_000;

    // Cells all modelled links of one sheet may cover together; keeps overlapping refs from multiplying a small file.
    internal const long MaxCellsPerSheet = 100_000;

    private const int MaxTooltipLength = 255;

    /// <summary>
    /// For each element in file order, what the model holds of it, or null when the element is kept as the
    /// file has it (damaged, too large, or past the sheet's budget). Where modelled elements overlap, the later wins.
    /// </summary>
    public static ModelledHyperlink?[] Classify(IReadOnlyList<HyperlinkElement> elements, IReadOnlyDictionary<string, Relationship> relationships)
    {
        var result = new ModelledHyperlink?[elements.Count];
        long budget = MaxCellsPerSheet;
        for (var i = 0; i < elements.Count; i++)
        {
            var element = elements[i];
            if (element.Ref is null
                || !ReferenceSyntax.TryParseA1Area(element.Ref, new CellAddress(1, 1), out var parsed)
                || parsed.Kind is not (AreaKind.Range or AreaKind.Cell))
                continue;
            var area = Area.Resolve(parsed, new CellAddress(1, 1));
            var cells = (long)(area.LastRow - area.FirstRow + 1) * (area.LastColumn - area.FirstColumn + 1);
            if (cells > MaxCellsPerElement || cells > budget)
                continue;

            string? address = null;
            if (element.RelationshipId is not null)
            {
                if (!relationships.TryGetValue(element.RelationshipId, out var relationship)
                    || !relationship.Is("hyperlink") || relationship.Target.Length == 0)
                    continue;
                address = relationship.Target;
            }

            var location = string.IsNullOrEmpty(element.Location) ? null : element.Location;
            string target;
            if (address is null)
            {
                if (location is null)
                    continue;
                target = "#" + location;
            }
            else
            {
                target = location is null ? address : address + "#" + location;
            }

            // A ScreenTip over Excel's limit is dropped; the link itself still works. Characters XML
            // cannot hold are escaped (_x0001_), as in cell text.
            var tooltip = element.Tooltip is { } raw && XmlText.Decode(raw) is { Length: > 0 and <= MaxTooltipLength } text ? text : null;
            budget -= cells;
            result[i] = new ModelledHyperlink(area, new Hyperlink(target, tooltip));
        }

        return result;
    }
}
