using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace SharpCell.Xlsx;

/// <summary>
/// The changes a save makes to the package: parts written again or added, parts removed, and the
/// content types and relationships that go with them. The writers of each part share one edit,
/// so [Content_Types].xml and each relationships part are changed once, however many writers add to them.
/// </summary>
internal sealed class PackageEdit(Package package)
{
    private static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    public const string OfficeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    private XDocument? _contentTypes;
    private readonly Dictionary<string, XDocument> _relationships = new(StringComparer.OrdinalIgnoreCase);

    public Package Package { get; } = package;

    /// <summary>Parts written again or added, by name without a leading '/'.</summary>
    public Dictionary<string, byte[]> Parts { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Parts of the source package the saved file leaves out.</summary>
    public HashSet<string> Removed { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Whether the part exists in the saved file: in the source and not removed, or added.</summary>
    public bool Exists(string part) => Parts.ContainsKey(part) || (Package.Exists(part) && !Removed.Contains(part));

    /// <summary>Adds a part with its content type and a relationship to it from <paramref name="source"/>.</summary>
    /// <returns>The relationship id.</returns>
    public string AddPart(string source, string part, byte[] bytes, string relationshipType, string contentType)
    {
        Parts[part] = bytes;
        Removed.Remove(part);
        SetContentType(part, contentType);
        return AddRelationship(source, relationshipType, part);
    }

    /// <summary>A part name not taken yet: <c>xl/worksheets/sheet4.xml</c> for the pattern <c>xl/worksheets/sheet{0}.xml</c>.</summary>
    public string FreeName(string pattern)
    {
        for (var i = 1; ; i++)
        {
            var name = string.Format(CultureInfo.InvariantCulture, pattern, i);
            if (!Exists(name))
                return name;
        }
    }

    public void SetContentType(string part, string contentType)
    {
        var root = ContentTypesDocument().Root!;
        root.Elements(ContentTypes + "Override")
            .Where(o => string.Equals(((string?)o.Attribute("PartName"))?.TrimStart('/'), part, StringComparison.OrdinalIgnoreCase))
            .Remove();
        root.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + part), new XAttribute("ContentType", contentType)));
    }

    /// <summary>
    /// Adds a relationship from <paramref name="source"/> to a part, or with <paramref name="external"/> to an
    /// address outside the package (a web address, a file), which is written as given.
    /// </summary>
    /// <returns>The new relationship's id, unique in the source part's relationships.</returns>
    public string AddRelationship(string source, string relationshipType, string target, bool external = false)
    {
        var document = RelationshipsDocument(source);
        var ids = document.Root!.Elements(Relationships + "Relationship").Select(r => (string?)r.Attribute("Id")).ToHashSet(StringComparer.Ordinal);
        var next = 1;
        string id;
        do
            id = "rIdSC" + next++.ToString(CultureInfo.InvariantCulture);
        while (ids.Contains(id));
        document.Root.Add(new XElement(Relationships + "Relationship", new XAttribute("Id", id),
            new XAttribute("Type", relationshipType), new XAttribute("Target", external ? target : RelativeTarget(source, target)),
            external ? new XAttribute("TargetMode", "External") : null));
        return id;
    }

    /// <summary>Removes one relationship of <paramref name="source"/>, leaving its target alone.</summary>
    public void RemoveRelationship(string source, string id) =>
        RelationshipsDocument(source).Root!.Elements(Relationships + "Relationship").Where(r => (string?)r.Attribute("Id") == id).Remove();

    /// <summary>Removes a part, its content type override and the relationships to it from <paramref name="source"/>.</summary>
    public void RemovePart(string source, string part)
    {
        Parts.Remove(part);
        if (Package.Exists(part))
            Removed.Add(part);
        ContentTypesDocument().Root!.Elements(ContentTypes + "Override")
            .Where(o => string.Equals(((string?)o.Attribute("PartName"))?.TrimStart('/'), part, StringComparison.OrdinalIgnoreCase))
            .Remove();
        var relationships = RelationshipsDocument(source);
        var folder = Folder(source);
        relationships.Root!.Elements(Relationships + "Relationship")
            .Where(r => (string?)r.Attribute("TargetMode") != "External"
                && string.Equals(Resolve(folder, (string?)r.Attribute("Target") ?? ""), part, StringComparison.OrdinalIgnoreCase))
            .Remove();
        var relationshipsPart = Package.RelationshipsPart(part);
        if (Package.Exists(relationshipsPart))
            Removed.Add(relationshipsPart);
    }

    /// <summary>
    /// Removes a part with every part that only it leads to through relationships: a sheet with its
    /// drawings, charts, comments and tables. Parts still reached from the package's root another
    /// way stay, such as a theme or an image another sheet shows too.
    /// </summary>
    public void RemoveTree(string source, string part)
    {
        var kept = Reachable("", part);
        var gone = Reachable(part, null);
        gone.ExceptWith(kept);
        RemovePart(source, part);
        foreach (var other in gone)
        {
            if (string.Equals(other, part, StringComparison.OrdinalIgnoreCase))
                continue;
            Parts.Remove(other);
            if (Package.Exists(other))
                Removed.Add(other);
            ContentTypesDocument().Root!.Elements(ContentTypes + "Override")
                .Where(o => string.Equals(((string?)o.Attribute("PartName"))?.TrimStart('/'), other, StringComparison.OrdinalIgnoreCase))
                .Remove();
            var relationshipsPart = Package.RelationshipsPart(other);
            _relationships.Remove(relationshipsPart);
            if (Package.Exists(relationshipsPart))
                Removed.Add(relationshipsPart);
        }
    }

    // The parts reached from a part (the root is "") through internal relationships, not passing through avoided.
    private HashSet<string> Reachable(string from, string? avoided)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var queue = new Queue<string>();
        queue.Enqueue(from);
        if (from.Length > 0)
            seen.Add(from);
        while (queue.Count > 0)
        {
            var part = queue.Dequeue();
            var relationshipsPart = part.Length == 0 ? "_rels/.rels" : Package.RelationshipsPart(part);
            var document = _relationships.TryGetValue(relationshipsPart, out var edited) ? edited : ReadXml(relationshipsPart);
            if (document?.Root is null)
                continue;
            var folder = Folder(part);
            foreach (var relationship in document.Root.Elements(Relationships + "Relationship"))
            {
                if ((string?)relationship.Attribute("TargetMode") == "External")
                    continue;
                var target = Resolve(folder, (string?)relationship.Attribute("Target") ?? "");
                if (target.Length == 0 || string.Equals(target, avoided, StringComparison.OrdinalIgnoreCase) || !Exists(target) || !seen.Add(target))
                    continue;
                queue.Enqueue(target);
            }
        }

        return seen;
    }

    /// <summary>Writes the changed content types and relationships into <see cref="Parts"/>.</summary>
    public void Complete()
    {
        if (_contentTypes is not null)
            Parts["[Content_Types].xml"] = XmlParts.Save(_contentTypes);
        foreach (var (part, document) in _relationships)
            Parts[part] = XmlParts.Save(document);
    }

    /// <summary>A part as XML: the version written in this edit, else the source's.</summary>
    public XDocument? ReadXml(string part)
    {
        if (Parts.TryGetValue(part, out var bytes))
            return XDocument.Load(new System.IO.MemoryStream(bytes));
        if (!Package.Exists(part) || Removed.Contains(part))
            return null;
        using var reader = Package.OpenXml(part);
        return XDocument.Load(reader);
    }

    private XDocument ContentTypesDocument() =>
        _contentTypes ??= ReadXml("[Content_Types].xml") ?? new XDocument(new XElement(ContentTypes + "Types"));

    private XDocument RelationshipsDocument(string source)
    {
        var part = Package.RelationshipsPart(source);
        if (!_relationships.TryGetValue(part, out var document))
        {
            document = ReadXml(part) ?? new XDocument(new XElement(Relationships + "Relationships"));
            _relationships[part] = document;
        }

        return document;
    }

    private static string Folder(string part)
    {
        var slash = part.LastIndexOf('/');
        return slash < 0 ? "" : part[..(slash + 1)];
    }

    // Targets are written relative to the source part's folder when the target is inside it.
    private static string RelativeTarget(string source, string target)
    {
        var folder = Folder(source);
        return target.StartsWith(folder, StringComparison.OrdinalIgnoreCase) ? target[folder.Length..] : "/" + target;
    }

    private static string Resolve(string folder, string target)
    {
        target = Uri.UnescapeDataString(target.Replace('\\', '/'));
        var path = target.StartsWith('/') ? target[1..] : folder + target;
        var segments = new List<string>();
        foreach (var segment in path.Split('/'))
        {
            if (segment is "" or ".")
                continue;
            if (segment == "..")
            {
                if (segments.Count > 0)
                    segments.RemoveAt(segments.Count - 1);
                continue;
            }

            segments.Add(segment);
        }

        return string.Join('/', segments);
    }
}
