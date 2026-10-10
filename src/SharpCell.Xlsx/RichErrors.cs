using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using SharpCell.Evaluation;

namespace SharpCell.Xlsx;

/// <summary>
/// The rich values behind #SPILL! and #CALC!: such a cell holds #VALUE! and its vm attribute points,
/// through xl/metadata.xml (valueMetadata → futureMetadata XLRICHVALUE), at a value in
/// xl/richData/rdrichvalue.xml whose structure is _error. Finds the value a cell needs or adds it,
/// creating the parts when the file has none. The shapes are the ones Excel writes.
/// </summary>
internal sealed class RichErrors
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RichData = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
    private static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string RichValueExtension = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}";
    private const string RichValueType = "XLRICHVALUE";
    private const string OfficeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string RichRelationships = "http://schemas.microsoft.com/office/2017/06/relationships/";

    // CT_Metadata's children in schema order.
    private static readonly string[] MetadataOrder =
        ["metadataTypes", "metadataStrings", "mdxMetadata", "futureMetadata", "cellMetadata", "valueMetadata", "extLst"];

    // rdRichValueTypes.xml as Excel writes it: which keys calculation ignores.
    private const string TypesXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
        + "<rvTypesInfo xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><global><keyFlags>"
        + "<key name=\"_Self\"><flag name=\"ExcludeFromFile\" value=\"1\"/><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>"
        + "<key name=\"_DisplayString\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>"
        + "<key name=\"_Flags\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>"
        + "<key name=\"_Format\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>"
        + "<key name=\"_SubLabel\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>"
        + "<key name=\"_Attribution\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>"
        + "<key name=\"_Icon\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>"
        + "<key name=\"_Display\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>"
        + "<key name=\"_CanonicalPropertyNames\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>"
        + "<key name=\"_ClassificationId\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>"
        + "</keyFlags></global></rvTypesInfo>";

    private readonly Package _package;
    private readonly string _workbookPart;
    private readonly string _folder;
    private readonly string? _typesPart;
    private string? _metadataPart;
    private string? _valuesPart;
    private string? _structuresPart;
    private XDocument? _metadata;
    private XDocument? _values;
    private XDocument? _structures;
    private readonly Dictionary<string, int> _vmByKey = new(StringComparer.Ordinal);
    private bool _changed;

    private RichErrors(Package package, string workbookPart, string? metadata, string? values, string? structures, string? types)
    {
        _package = package;
        _workbookPart = workbookPart;
        var slash = workbookPart.LastIndexOf('/');
        _folder = slash < 0 ? "" : workbookPart[..(slash + 1)];
        _metadataPart = metadata;
        _valuesPart = values;
        _structuresPart = structures;
        _typesPart = types;
    }

    /// <summary>Finds the parts; they are read only when a cell needs a rich error.</summary>
    public static RichErrors Open(Package package, string workbookPart, Dictionary<string, Relationship> relationships)
    {
        // As the reader does: files without the relationships keep the parts at their usual names.
        string? Find(string kind, string usual) =>
            WorkbookReader.FirstOfType(relationships, kind)?.Target is { } target && package.Exists(target) ? target
            : package.Exists(usual) ? usual : null;

        return new RichErrors(package, workbookPart,
            Find("sheetMetadata", "xl/metadata.xml"),
            Find("rdRichValue", "xl/richData/rdrichvalue.xml"),
            Find("rdRichValueStructure", "xl/richData/rdrichvaluestructure.xml"),
            Find("rdRichValueTypes", "xl/richData/rdRichValueTypes.xml"));
    }

    /// <summary>The 1-based vm attribute for a cell showing the error.</summary>
    /// <param name="kind">#SPILL! or #CALC!.</param>
    /// <param name="intendedSpill">For a #SPILL! anchor that cells block, the area it wanted; otherwise null.</param>
    /// <param name="propagated">Whether the cell passes on an error it read from another cell.</param>
    public int VmOf(ErrorKind kind, Area? intendedSpill, bool propagated)
    {
        Load();
        var entry = Entry(kind, intendedSpill, propagated);
        var key = Key(entry.Select(e => e.Key), entry.Select(e => e.Value.ToString(CultureInfo.InvariantCulture)));
        if (_vmByKey.TryGetValue(key, out var vm))
            return vm;

        _changed = true;
        var structure = StructureIndex(entry.Select(e => (e.Key, e.Type)).ToArray());
        var value = Append(_values!.Root!, new XElement(RichData + "rv", new XAttribute("s", structure),
            entry.Select(e => new XElement(RichData + "v", e.Value.ToString(CultureInfo.InvariantCulture)))));

        var metadata = _metadata!.Root!;
        if (metadata.Attribute(XNamespace.Xmlns + "xlrd") is null)
            metadata.Add(new XAttribute(XNamespace.Xmlns + "xlrd", RichData));
        var type = TypeIndex();
        var future = metadata.Elements(Main + "futureMetadata").FirstOrDefault(e => (string?)e.Attribute("name") == RichValueType)
            ?? Insert(metadata, new XElement(Main + "futureMetadata", new XAttribute("name", RichValueType), new XAttribute("count", 0)));
        var futureIndex = Append(future, new XElement(Main + "bk", new XElement(Main + "extLst",
            new XElement(Main + "ext", new XAttribute("uri", RichValueExtension), new XElement(RichData + "rvb", new XAttribute("i", value))))));
        var valueMetadata = metadata.Element(Main + "valueMetadata") ?? Insert(metadata, new XElement(Main + "valueMetadata", new XAttribute("count", 0)));
        vm = Append(valueMetadata, new XElement(Main + "bk", new XElement(Main + "rc", new XAttribute("t", type), new XAttribute("v", futureIndex)))) + 1;
        _vmByKey[key] = vm;
        return vm;
    }

    /// <summary>Adds the changed and created parts, with their content types and relationships.</summary>
    public void Save(Dictionary<string, byte[]> parts)
    {
        if (!_changed)
            return;
        var created = new List<(string Part, string Relationship, string ContentType)>();
        Put(parts, created, ref _metadataPart, _metadata!, "metadata.xml",
            OfficeRelationships + "sheetMetadata", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml");
        Put(parts, created, ref _valuesPart, _values!, "richData/rdrichvalue.xml",
            RichRelationships + "rdRichValue", "application/vnd.ms-excel.rdrichvalue+xml");
        Put(parts, created, ref _structuresPart, _structures!, "richData/rdrichvaluestructure.xml",
            RichRelationships + "rdRichValueStructure", "application/vnd.ms-excel.rdrichvaluestructure+xml");
        if (_typesPart is null)
        {
            var part = _folder + "richData/rdRichValueTypes.xml";
            parts[part] = Encoding.UTF8.GetBytes(TypesXml);
            created.Add((part, RichRelationships + "rdRichValueTypes", "application/vnd.ms-excel.rdrichvaluetypes+xml"));
        }

        if (created.Count == 0)
            return;
        parts["[Content_Types].xml"] = WithOverrides(created);
        var rels = Package.RelationshipsPart(_workbookPart);
        parts[rels] = WithRelationships(rels, created);
    }

    // Keys in the order Excel writes them, with their values and types (i: integer, b: boolean).
    private static (string Key, int Value, string Type)[] Entry(ErrorKind kind, Area? intendedSpill, bool propagated)
    {
        var errorType = (int)kind - 1;   // the ERROR.TYPE code minus one, as RichValues reads it
        if (kind == ErrorKind.Spill && intendedSpill is { } area)
            return [("colOffset", area.Columns - 1, "i"), ("errorType", errorType, "i"), ("rwOffset", area.Rows - 1, "i"), ("subType", 1, "i")];
        return propagated
            ? [("errorType", errorType, "i"), ("propagated", 1, "b")]
            : [("errorType", errorType, "i"), ("subType", 0, "i")];
    }

    private static string Key(IEnumerable<string> keys, IEnumerable<string> values) =>
        string.Join(';', keys.Zip(values, (k, v) => k + "=" + v));

    private void Load()
    {
        if (_metadata is not null)
            return;
        _metadata = Read(_metadataPart) ?? new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(Main + "metadata", new XAttribute(XNamespace.Xmlns + "xlrd", RichData)));
        _values = Read(_valuesPart) ?? new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(RichData + "rvData", new XAttribute("count", 0)));
        _structures = Read(_structuresPart) ?? new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(RichData + "rvStructures", new XAttribute("count", 0)));

        // The rich errors the file already has, by content, so equal errors share one value.
        var types = _metadata.Root!.Element(Main + "metadataTypes")?.Elements(Main + "metadataType")
            .Select(e => (string?)e.Attribute("name")).ToList() ?? [];
        var future = _metadata.Root.Elements(Main + "futureMetadata").FirstOrDefault(e => (string?)e.Attribute("name") == RichValueType)?
            .Elements(Main + "bk").Select(bk => (int?)bk.Descendants(RichData + "rvb").FirstOrDefault()?.Attribute("i") ?? -1).ToList() ?? [];
        var structures = _structures.Root!.Elements(RichData + "s").ToList();
        var values = _values.Root!.Elements(RichData + "rv").ToList();
        var blocks = _metadata.Root.Element(Main + "valueMetadata")?.Elements(Main + "bk").ToList() ?? [];
        for (var i = 0; i < blocks.Count; i++)
        {
            var rc = blocks[i].Element(Main + "rc");
            var type = (int?)rc?.Attribute("t") ?? 0;
            var index = (int?)rc?.Attribute("v") ?? -1;
            if (type < 1 || type > types.Count || types[type - 1] != RichValueType || index < 0 || index >= future.Count)
                continue;
            var rv = future[index];
            if (rv < 0 || rv >= values.Count)
                continue;
            var s = (int?)values[rv].Attribute("s") ?? -1;
            if (s < 0 || s >= structures.Count || (string?)structures[s].Attribute("t") != "_error")
                continue;
            var key = Key(structures[s].Elements(RichData + "k").Select(k => (string?)k.Attribute("n") ?? ""),
                values[rv].Elements(RichData + "v").Select(v => v.Value));
            _vmByKey.TryAdd(key, i + 1);
        }
    }

    private XDocument? Read(string? part)
    {
        if (part is null)
            return null;
        using var reader = _package.OpenXml(part);
        return XDocument.Load(reader);
    }

    // The index of an _error structure with these keys, added if the file has none.
    private int StructureIndex((string Key, string Type)[] keys)
    {
        var structures = _structures!.Root!.Elements(RichData + "s").ToList();
        for (var i = 0; i < structures.Count; i++)
        {
            if ((string?)structures[i].Attribute("t") == "_error"
                && structures[i].Elements(RichData + "k").Select(k => (string?)k.Attribute("n")).SequenceEqual(keys.Select(k => k.Key)))
                return i;
        }

        return Append(_structures.Root!, new XElement(RichData + "s", new XAttribute("t", "_error"),
            keys.Select(k => new XElement(RichData + "k", new XAttribute("n", k.Key), new XAttribute("t", k.Type)))));
    }

    // The 1-based index of the XLRICHVALUE metadata type, added if the file has none.
    private int TypeIndex()
    {
        var root = _metadata!.Root!;
        var section = root.Element(Main + "metadataTypes") ?? Insert(root, new XElement(Main + "metadataTypes", new XAttribute("count", 0)));
        var types = section.Elements(Main + "metadataType").ToList();
        var found = types.FindIndex(t => (string?)t.Attribute("name") == RichValueType);
        if (found >= 0)
            return found + 1;
        return Append(section, new XElement(Main + "metadataType", new XAttribute("name", RichValueType),
            new XAttribute("minSupportedVersion", 120000), new XAttribute("copy", 1), new XAttribute("pasteAll", 1),
            new XAttribute("pasteValues", 1), new XAttribute("merge", 1), new XAttribute("splitFirst", 1),
            new XAttribute("rowColShift", 1), new XAttribute("clearFormats", 1), new XAttribute("clearComments", 1),
            new XAttribute("assign", 1), new XAttribute("coerce", 1))) + 1;
    }

    // Adds a child, sets the parent's count to the number of such children, and returns its 0-based index.
    private static int Append(XElement parent, XElement child)
    {
        parent.Add(child);
        var count = parent.Elements(child.Name).Count();
        parent.SetAttributeValue("count", count);
        return count - 1;
    }

    // Puts a metadata section in schema order: after the last section that may precede it.
    private static XElement Insert(XElement root, XElement section)
    {
        var rank = Array.IndexOf(MetadataOrder, section.Name.LocalName);
        var previous = root.Elements().LastOrDefault(e => e.Name.Namespace == Main
            && Array.IndexOf(MetadataOrder, e.Name.LocalName) is var r && r >= 0 && r <= rank);
        if (previous is null)
            root.AddFirst(section);
        else
            previous.AddAfterSelf(section);
        return section;
    }

    private void Put(Dictionary<string, byte[]> parts, List<(string Part, string Relationship, string ContentType)> created,
        ref string? part, XDocument document, string name, string relationship, string contentType)
    {
        if (part is null)
        {
            part = _folder + name;
            created.Add((part, relationship, contentType));
        }

        parts[part] = XmlParts.Save(document);
    }

    private byte[] WithOverrides(List<(string Part, string Relationship, string ContentType)> created)
    {
        var document = Read("[Content_Types].xml")!;
        foreach (var (part, _, contentType) in created)
            document.Root!.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + part), new XAttribute("ContentType", contentType)));
        return XmlParts.Save(document);
    }

    private byte[] WithRelationships(string rels, List<(string Part, string Relationship, string ContentType)> created)
    {
        var document = _package.Exists(rels) ? Read(rels)! : new XDocument(new XElement(PackageRelationships + "Relationships"));
        var ids = document.Root!.Elements(PackageRelationships + "Relationship").Select(r => (string?)r.Attribute("Id")).ToHashSet(StringComparer.Ordinal);
        var next = 1;
        foreach (var (part, relationship, _) in created)
        {
            string id;
            do
                id = "rIdSC" + next++.ToString(CultureInfo.InvariantCulture);
            while (ids.Contains(id));
            document.Root.Add(new XElement(PackageRelationships + "Relationship", new XAttribute("Id", id),
                new XAttribute("Type", relationship), new XAttribute("Target", part[_folder.Length..])));
        }

        return XmlParts.Save(document);
    }
}
