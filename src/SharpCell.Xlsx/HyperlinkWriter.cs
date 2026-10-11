using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace SharpCell.Xlsx;

/// <summary>
/// The links of a sheet whose &lt;hyperlinks&gt; element is written again: the file's elements the model
/// does not hold, as the file has them, and the relationship id of each address the model links to.
/// </summary>
internal sealed record SheetHyperlinks(IReadOnlyList<XElement> Kept, IReadOnlyDictionary<string, string> Ids);

/// <summary>
/// Writes a sheet's hyperlinks: one &lt;hyperlink&gt; per cell, its address (a web or e-mail address, a
/// file) in an external relationship of the sheet and a place in the workbook, or the part after '#'
/// of an address, in location.
/// </summary>
internal static class HyperlinkWriter
{
    public const string RelationshipType = PackageEdit.OfficeRelationships + "hyperlink";

    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>What a save writes for the sheet's links, with their relationships prepared; null when the file's element stays.</summary>
    public static SheetHyperlinks? Plan(PackageEdit edit, SheetSource source, string sheetPart, Worksheet sheet, SheetEditMap map)
    {
        // A range of the file grows with rows inserted inside it and a cell's link does not, so
        // links the file gave a sheet whose rows or columns moved are written from the model.
        if (!source.Look.Mapped(map, sheet).HyperlinksChanged(sheet) && !(map.HasEdits(sheet) && source.Look.Hyperlinks.Count > 0))
            return null;
        var kept = source.Part is { } part ? Kept(edit.Package, part) : [];
        var ids = Prepare(edit, sheetPart, sheet, [.. kept.Select(k => k.RelationshipId).OfType<string>()]);
        return new SheetHyperlinks([.. kept.Select(k => k.Element)], ids);
    }

    /// <summary>
    /// The relationship id of each address the sheet's links use (address → id). An existing hyperlink
    /// relationship with the same address is used again; one that neither the model nor a kept element
    /// uses any more is removed.
    /// </summary>
    public static IReadOnlyDictionary<string, string> Prepare(PackageEdit edit, string sheetPart, Worksheet sheet, IReadOnlyCollection<string> keptIds)
    {
        var existing = edit.Package.ReadRelationships(sheetPart, includeExternal: true).Values.Where(r => r.Is("hyperlink")).ToList();
        var known = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var relationship in existing.Where(r => r.IsExternal))
            known.TryAdd(relationship.Target, relationship.Id);

        var ids = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (_, link) in Ordered(sheet))
        {
            if (Split(link.Target).Address is not { } address || ids.ContainsKey(address))
                continue;
            ids[address] = known.TryGetValue(address, out var id) ? id : edit.AddRelationship(sheetPart, RelationshipType, address, external: true);
        }

        var used = new HashSet<string>(ids.Values, StringComparer.Ordinal);
        used.UnionWith(keptIds);
        foreach (var relationship in existing.Where(r => !used.Contains(r.Id)))
            edit.RemoveRelationship(sheetPart, relationship.Id);
        return ids;
    }

    /// <summary>The &lt;hyperlinks&gt; element: the kept elements first, then the model's links in row, then column order; null when there are none.</summary>
    public static XElement? Element(XNamespace ns, Worksheet sheet, IReadOnlyDictionary<string, string> ids, IEnumerable<XElement> kept)
    {
        var element = new XElement(ns + "hyperlinks", kept);
        var r = element.Elements().SelectMany(e => e.Attributes()).FirstOrDefault(a => a.Name.LocalName == "id" && a.Name.Namespace != XNamespace.None)?.Name.Namespace
            ?? Relationships;

        // A model's link written after a kept element covering the same cell wins, as on reading.
        foreach (var (cell, link) in Ordered(sheet))
        {
            var (address, location) = Split(link.Target);
            if (address is null && location is null)
                continue;   // "#" alone goes nowhere
            element.Add(new XElement(ns + "hyperlink",
                new XAttribute("ref", cell.ToString()),
                address is null ? null : new XAttribute(r + "id", ids[address]),
                location is null ? null : new XAttribute("location", location),
                link.Tooltip is null ? null : new XAttribute("tooltip", link.Tooltip)));
        }

        if (!element.HasElements)
            return null;
        if (element.Descendants().Attributes().Any(a => a.Name.Namespace == r))
            element.SetAttributeValue(XNamespace.Xmlns + "r", r.NamespaceName);
        return element;
    }

    // The file's <hyperlink> elements the model does not hold, in file order, with their relationship ids.
    private static List<(XElement Element, string? RelationshipId)> Kept(Package package, string part)
    {
        var elements = new List<XElement>();
        var read = new List<HyperlinkElement>();
        using (var reader = package.OpenXml(part))
        {
            reader.MoveToContent();
            var ns = reader.NamespaceURI;
            reader.Read();
            while (!reader.EOF)
            {
                if (reader.NodeType != XmlNodeType.Element)
                    reader.Read();
                else if (reader.Depth == 1 && (reader.LocalName != "hyperlinks" || reader.NamespaceURI != ns))
                    reader.Skip();
                else if (reader.Depth == 2 && reader.LocalName == "hyperlink")
                {
                    read.Add(HyperlinkElement.Read(reader));
                    elements.Add((XElement)XNode.ReadFrom(reader));
                }
                else
                    reader.Read();
            }
        }

        if (elements.Count == 0)
            return [];
        var relationships = read.Exists(link => link.RelationshipId is not null)
            ? package.ReadRelationships(part, includeExternal: true)
            : new Dictionary<string, Relationship>();
        var modelled = HyperlinkElements.Classify(read, relationships);
        return [.. elements.Select((element, i) => (element, read[i].RelationshipId)).Where((_, i) => modelled[i] is null)];
    }

    // The address goes to the relationship; what follows '#' is a place in the target (or, alone, in the workbook).
    private static (string? Address, string? Location) Split(string target)
    {
        var hash = target.IndexOf('#', StringComparison.Ordinal);
        if (hash < 0)
            return (target, null);
        var location = hash == target.Length - 1 ? null : target[(hash + 1)..];
        return (hash == 0 ? null : target[..hash], location);
    }

    private static IEnumerable<KeyValuePair<CellAddress, Hyperlink>> Ordered(Worksheet sheet) =>
        sheet.Hyperlinks.OrderBy(link => link.Key.Row).ThenBy(link => link.Key.Column);
}
