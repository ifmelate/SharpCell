using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace SharpCell.Xlsx;

/// <summary>
/// Notes in a file: their text and authors in a comments part, and a box for each in the sheet's
/// legacy VML drawing, which Excel needs to show them. A threaded comment also has an entry in both;
/// those are not notes and are kept as they are.
/// </summary>
internal static partial class NoteParts
{
    public const string CommentsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
    public const string VmlContentType = "application/vnd.openxmlformats-officedocument.vmlDrawing";

    private static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace Office = "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace Excel = "urn:schemas-microsoft-com:office:excel";
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// Reads the notes of a sheet into it. Like links, notes only decorate cells: a damaged part is
    /// skipped, never a reason to refuse the file.
    /// </summary>
    public static void Read(Package package, string sheetPart, Worksheet sheet)
    {
        var relationships = package.ReadRelationships(sheetPart);
        if (relationships.Values.FirstOrDefault(r => r.Is("comments") && package.Exists(r.Target)) is not { } comments)
            return;
        try
        {
            var threaded = ThreadedCells(relationships.Values.Select(r => r.Target).Where(package.Exists)
                .Where(target => relationships.Values.Any(r => r.Target == target && r.Is("threadedComment"))), package.ReadXmlDocument);
            var visible = new HashSet<CellAddress>();
            foreach (var vml in relationships.Values.Where(r => r.Is("vmlDrawing") && package.Exists(r.Target)))
            {
                if (ReadVml(package.ReadText(vml.Target)) is { } drawing)
                {
                    foreach (var (cell, shape) in NoteShapes(drawing))
                    {
                        if (IsVisible(shape))
                            visible.Add(cell);
                    }
                }
            }

            var root = package.ReadXmlDocument(comments.Target)?.Root;
            if (root is null)
                return;
            var authors = root.Element(Main + "authors")?.Elements(Main + "author").Select(a => a.Value).ToList() ?? [];
            foreach (var comment in root.Element(Main + "commentList")?.Elements(Main + "comment") ?? [])
            {
                if ((string?)comment.Attribute("ref") is not { } reference || !CellAddress.TryParse(reference, out var cell) || threaded.Contains(cell))
                    continue;
                var authorId = int.TryParse((string?)comment.Attribute("authorId"), NumberStyles.None, CultureInfo.InvariantCulture, out var id) ? id : -1;
                var author = authorId >= 0 && authorId < authors.Count && authors[authorId].Length > 0 ? authors[authorId] : null;
                var text = Text(comment.Element(Main + "text"));
                if (text.Length > 32767)
                    continue;
                sheet.SetNote(cell.Row, cell.Column, new Note(text, author) { IsVisible = visible.Contains(cell) });
            }
        }
        catch (XmlException)
        {
        }
    }

    // A comment's text: its runs (or its one t) without phonetic guides.
    private static string Text(XElement? text)
    {
        if (text is null)
            return "";
        var builder = new StringBuilder();
        foreach (var element in text.Elements())
        {
            if (element.Name == Main + "t")
                builder.Append(element.Value);
            else if (element.Name == Main + "r")
                builder.Append(element.Element(Main + "t")?.Value);
        }

        return builder.ToString().Replace("\r\n", "\n", StringComparison.Ordinal);
    }

    private static HashSet<CellAddress> ThreadedCells(IEnumerable<string> parts, Func<string, XDocument?> read)
    {
        var cells = new HashSet<CellAddress>();
        foreach (var part in parts)
        {
            foreach (var comment in read(part)?.Root?.Elements().Where(e => e.Name.LocalName == "threadedComment") ?? [])
            {
                if ((string?)comment.Attribute("ref") is { } reference && CellAddress.TryParse(reference, out var cell))
                    cells.Add(cell);
            }
        }

        return cells;
    }

    /// <summary>
    /// A VML drawing as XML. Excel's own is well-formed except, at times, for HTML line breaks
    /// (<c>&lt;br&gt;</c>) in a box's text; null when it still cannot be read.
    /// </summary>
    private static XDocument? ReadVml(string text)
    {
        try
        {
            return XDocument.Parse(text, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException)
        {
        }

        try
        {
            return XDocument.Parse(LineBreak().Replace(text, "<br/>"), LoadOptions.PreserveWhitespace);
        }
        catch (XmlException)
        {
            return null;
        }
    }

    [GeneratedRegex(@"<br\s*>", RegexOptions.IgnoreCase)]
    private static partial Regex LineBreak();

    // The note boxes of a drawing by their cells: shapes whose client data is a Note.
    private static IEnumerable<(CellAddress Cell, XElement Shape)> NoteShapes(XDocument drawing)
    {
        foreach (var data in drawing.Descendants(Excel + "ClientData"))
        {
            if ((string?)data.Attribute("ObjectType") != "Note" || data.Parent is not { } shape)
                continue;
            if (int.TryParse(data.Element(Excel + "Row")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var row)
                && int.TryParse(data.Element(Excel + "Column")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var column)
                && row < CellAddress.MaxRow && column < CellAddress.MaxColumn)
                yield return (new CellAddress(row + 1, column + 1), shape);
        }
    }

    private static bool IsVisible(XElement shape) => shape.Element(Excel + "ClientData")?.Element(Excel + "Visible") is { } visible
        && !string.Equals(visible.Value.Trim(), "False", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// What saving a sheet's notes takes, decided before the worksheet part is written: which parts
    /// hold them, and whether the worksheet must gain or lose its legacyDrawing element.
    /// </summary>
    /// <param name="Sheet">The sheet.</param>
    /// <param name="SheetPart">Its worksheet part.</param>
    /// <param name="Comments">The comments part, new or the file's; null when the sheet keeps none.</param>
    /// <param name="Drawing">The VML part with the boxes, new or the file's; null when the sheet keeps none.</param>
    /// <param name="NewDrawing">Whether the VML part is new.</param>
    /// <param name="LegacyDrawingId">The relationship id the worksheet's new legacyDrawing element names, or null.</param>
    /// <param name="RemoveLegacyDrawing">Whether the worksheet's legacyDrawing element goes.</param>
    /// <param name="ThreadedCells">Cells with threaded comments, whose entries are kept, where the edits put them.</param>
    public sealed record Plan(Worksheet Sheet, string SheetPart, string? Comments, string? Drawing, bool NewDrawing,
        string? LegacyDrawingId, bool RemoveLegacyDrawing, IReadOnlySet<CellAddress> ThreadedCells);

    /// <returns>The plan, or null when the sheet's notes are as the file has them.</returns>
    public static Plan? Prepare(PackageEdit edit, SheetSource source, string sheetPart, Worksheet sheet, SheetEditMap map)
    {
        if (!source.Look.Mapped(map, sheet).NotesChanged(sheet))
            return null;
        var package = edit.Package;
        var relationships = source.Part is { } part ? package.ReadRelationships(part) : [];
        var threaded = new HashSet<CellAddress>();
        foreach (var cell in ThreadedCells(relationships.Values.Where(r => r.Is("threadedComment") && package.Exists(r.Target)).Select(r => r.Target), package.ReadXmlDocument))
        {
            if (map.Map(sheet, cell) is { } moved)
                threaded.Add(moved);
        }

        var comments = relationships.Values.FirstOrDefault(r => r.Is("comments") && package.Exists(r.Target))?.Target;
        var legacyId = source.Part is { } sheetSource ? LegacyDrawingId(package, sheetSource) : null;
        var drawing = legacyId is not null && relationships.TryGetValue(legacyId, out var vml) && vml.Is("vmlDrawing") && package.Exists(vml.Target)
            ? vml.Target
            : null;
        if (legacyId is not null && drawing is null)
            throw new NotSupportedException($"The legacy drawing of sheet '{sheet.Name}' is missing from the file; its notes cannot be saved.");

        var keep = sheet.NoteMap.Count > 0 || threaded.Count > 0;
        if (!keep)
        {
            var otherShapes = drawing is not null && ReadVml(package.ReadText(drawing)) is { } document
                && document.Descendants(Vml + "shape").Any(shape => (string?)shape.Element(Excel + "ClientData")?.Attribute("ObjectType") != "Note");
            return new Plan(sheet, sheetPart, comments, drawing, false, null, drawing is not null && !otherShapes, threaded);
        }

        if (comments is null)
        {
            comments = edit.FreeName("xl/comments{0}.xml");
            edit.Parts[comments] = [];   // claims the name
            edit.SetContentType(comments, CommentsContentType);
            edit.AddRelationship(sheetPart, PackageEdit.OfficeRelationships + "comments", comments);
        }

        if (drawing is not null)
            return new Plan(sheet, sheetPart, comments, drawing, false, null, false, threaded);
        drawing = edit.FreeName("xl/drawings/vmlDrawing{0}.vml");
        edit.Parts[drawing] = [];
        edit.SetContentType(drawing, VmlContentType);
        var id = edit.AddRelationship(sheetPart, PackageEdit.OfficeRelationships + "vmlDrawing", drawing);
        return new Plan(sheet, sheetPart, comments, drawing, true, id, false, threaded);
    }

    // The relationship id of the worksheet's legacyDrawing element, read from the part as the file has it.
    private static string? LegacyDrawingId(Package package, string part)
    {
        using var reader = package.OpenXml(part);
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1 && reader.LocalName == "legacyDrawing")
                return reader.GetAttribute("id", PackageEdit.OfficeRelationships.TrimEnd('/'));
        }

        return null;
    }

    /// <summary>
    /// Writes the comments and VML parts of the plans, after the other parts were moved for inserted
    /// and deleted rows: the file's entries are then where the model's notes are. A note that did not
    /// change keeps its entry, with its formatted text and its box; a new or changed one is written
    /// as plain text in a box beside its cell.
    /// </summary>
    public static void Write(PackageEdit edit, IEnumerable<Plan> plans)
    {
        var idMaps = UsedIdMaps(edit);
        foreach (var plan in plans)
        {
            var empty = plan.Sheet.NoteMap.Count == 0 && plan.ThreadedCells.Count == 0;
            if (plan.Comments is not null && empty)
                edit.RemovePart(plan.SheetPart, plan.Comments);
            else if (plan.Comments is not null)
                WriteComments(edit, plan);
            if (plan.Drawing is not null && plan.RemoveLegacyDrawing)
                edit.RemovePart(plan.SheetPart, plan.Drawing);
            else if (plan.Drawing is not null)
                WriteDrawing(edit, plan, idMaps);
        }
    }

    // The comments part, written again: entries of threaded comments and of notes whose text stayed
    // are kept with their formatting.
    private static void WriteComments(PackageEdit edit, Plan plan)
    {
        var notes = plan.Sheet.NoteMap;
        var part = plan.Comments!;
        var old = edit.Parts.TryGetValue(part, out var bytes) && bytes.Length == 0 ? null : edit.ReadXml(part)?.Root;
        var ns = old?.Name.Namespace ?? Main;
        var oldAuthors = old?.Element(ns + "authors")?.Elements(ns + "author").Select(a => a.Value).ToList() ?? [];
        var oldComments = new Dictionary<CellAddress, XElement>();
        foreach (var comment in old?.Element(ns + "commentList")?.Elements(ns + "comment") ?? [])
        {
            if ((string?)comment.Attribute("ref") is { } reference && CellAddress.TryParse(reference, out var cell))
                oldComments.TryAdd(cell, comment);
        }

        var authors = new List<string>();
        int AuthorId(string? author)
        {
            var index = authors.IndexOf(author ?? "");
            if (index >= 0)
                return index;
            authors.Add(author ?? "");
            return authors.Count - 1;
        }

        var entries = new SortedDictionary<CellAddress, XElement>(Comparer<CellAddress>.Create((a, b) => a.Row != b.Row ? a.Row.CompareTo(b.Row) : a.Column.CompareTo(b.Column)));
        foreach (var cell in plan.ThreadedCells)
        {
            if (oldComments.TryGetValue(cell, out var entry))
            {
                var author = int.TryParse((string?)entry.Attribute("authorId"), NumberStyles.None, CultureInfo.InvariantCulture, out var id) && id < oldAuthors.Count ? oldAuthors[id] : "";
                entry.SetAttributeValue("authorId", AuthorId(author));
                entries[cell] = entry;
            }
        }

        foreach (var (cell, note) in notes)
        {
            if (plan.ThreadedCells.Contains(cell))
                continue;
            XElement entry;
            if (oldComments.TryGetValue(cell, out var existing) && Text(existing.Element(ns + "text")) == note.Text)
            {
                entry = existing;
            }
            else
            {
                entry = new XElement(ns + "comment", new XAttribute("ref", cell.ToString()), new XAttribute("authorId", 0), new XAttribute("shapeId", 0),
                    new XElement(ns + "text", new XElement(ns + "r",
                        new XElement(ns + "rPr", new XElement(ns + "sz", new XAttribute("val", 9)), new XElement(ns + "color", new XAttribute("indexed", 81)),
                            new XElement(ns + "rFont", new XAttribute("val", "Tahoma")), new XElement(ns + "family", new XAttribute("val", 2))),
                        new XElement(ns + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), note.Text))));
            }

            entry.SetAttributeValue("authorId", AuthorId(note.Author));
            entries[cell] = entry;
        }

        var root = new XElement(ns + "comments", old?.Attributes().Where(a => a.IsNamespaceDeclaration),
            new XElement(ns + "authors", authors.Select(a => new XElement(ns + "author", a))),
            new XElement(ns + "commentList", entries.Values),
            old?.Element(ns + "extLst"));
        edit.Parts[part] = XmlParts.Save(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root));
    }

    // The VML part with a box for each note: boxes of notes that stayed are kept, those of removed
    // notes go, and new notes get one.
    private static void WriteDrawing(PackageEdit edit, Plan plan, HashSet<int> idMaps)
    {
        var drawing = plan.Drawing!;
        XDocument document;
        if (plan.NewDrawing)
        {
            var map = 1;
            while (!idMaps.Add(map))
                map++;
            document = new XDocument(new XElement("xml",
                new XAttribute(XNamespace.Xmlns + "v", Vml), new XAttribute(XNamespace.Xmlns + "o", Office), new XAttribute(XNamespace.Xmlns + "x", Excel),
                new XElement(Office + "shapelayout", new XAttribute(Vml + "ext", "edit"),
                    new XElement(Office + "idmap", new XAttribute(Vml + "ext", "edit"), new XAttribute("data", map)))));
        }
        else
        {
            var text = edit.Parts.TryGetValue(drawing, out var bytes) ? Encoding.UTF8.GetString(bytes) : edit.Package.ReadText(drawing);
            document = ReadVml(text)
                ?? throw new NotSupportedException($"The legacy drawing of sheet '{plan.Sheet.Name}' is not well-formed XML; its notes cannot be saved.");
        }

        var root = document.Root!;
        var notes = plan.Sheet.NoteMap;
        foreach (var (cell, shape) in NoteShapes(document).ToList())
        {
            if (plan.ThreadedCells.Contains(cell))
                continue;
            if (!notes.TryGetValue(cell, out var note))
            {
                shape.Remove();
                continue;
            }

            SetVisible(shape, note.IsVisible);
        }

        var boxed = NoteShapes(document).Select(s => s.Cell).ToHashSet();
        var missing = notes.Keys.Where(cell => !boxed.Contains(cell)).OrderBy(c => c.Row).ThenBy(c => c.Column).ToList();
        if (missing.Count > 0)
        {
            if (!root.Elements(Vml + "shapetype").Any(t => (string?)t.Attribute("id") == "_x0000_t202"))
                root.Add(NoteShapeType());
            var idMap = root.Element(Office + "shapelayout")?.Element(Office + "idmap") is { } idmap
                && int.TryParse(((string?)idmap.Attribute("data"))?.Split(',')[0], NumberStyles.None, CultureInfo.InvariantCulture, out var data) ? data : 1;
            var next = Math.Max(idMap * 1024, root.Descendants(Vml + "shape").Select(ShapeNumber).DefaultIfEmpty(idMap * 1024).Max()) + 1;
            foreach (var cell in missing)
                root.Add(NoteShape(next++, cell, notes[cell].IsVisible));
        }

        edit.Parts[drawing] = Encoding.UTF8.GetBytes(document.Root!.ToString(SaveOptions.DisableFormatting));
    }

    private static int ShapeNumber(XElement shape) =>
        (string?)shape.Attribute("id") is { } id && id.StartsWith("_x0000_s", StringComparison.Ordinal)
        && int.TryParse(id.AsSpan(8), NumberStyles.None, CultureInfo.InvariantCulture, out var number) ? number : 0;

    private static void SetVisible(XElement shape, bool visible)
    {
        var data = shape.Element(Excel + "ClientData")!;
        var style = ((string?)shape.Attribute("style") ?? "").Replace(";visibility:hidden", "", StringComparison.Ordinal).Replace("visibility:hidden", "", StringComparison.Ordinal);
        if (IsVisible(shape) == visible)
            return;
        data.Elements(Excel + "Visible").Remove();
        if (visible)
        {
            data.Add(new XElement(Excel + "Visible"));
            shape.SetAttributeValue("style", style);
        }
        else
        {
            shape.SetAttributeValue("style", style.Length == 0 ? "visibility:hidden" : style.TrimEnd(';') + ";visibility:hidden");
        }
    }

    // Excel's shape type for note boxes.
    private static XElement NoteShapeType() =>
        new(Vml + "shapetype", new XAttribute("id", "_x0000_t202"), new XAttribute("coordsize", "21600,21600"), new XAttribute(Office + "spt", "202"),
            new XAttribute("path", "m,l,21600r21600,l21600,xe"),
            new XElement(Vml + "stroke", new XAttribute("joinstyle", "miter")),
            new XElement(Vml + "path", new XAttribute("gradientshapeok", "t"), new XAttribute(Office + "connecttype", "rect")));

    // A box as Excel makes for a new note: to the right of its cell, two columns wide and four rows high.
    private static XElement NoteShape(int number, CellAddress cell, bool visible)
    {
        var (row, column) = (cell.Row - 1, cell.Column - 1);
        var top = Math.Max(row - 1, 0);
        var anchor = string.Join(", ", column + 1, 15, top, 10, column + 3, 15, top + 4, 4);
        return new XElement(Vml + "shape", new XAttribute("id", "_x0000_s" + number.ToString(CultureInfo.InvariantCulture)), new XAttribute("type", "#_x0000_t202"),
            new XAttribute("style", "position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:" + number.ToString(CultureInfo.InvariantCulture) + (visible ? "" : ";visibility:hidden")),
            new XAttribute("fillcolor", "#ffffe1"), new XAttribute(Office + "insetmode", "auto"),
            new XElement(Vml + "fill", new XAttribute("color2", "#ffffe1")),
            new XElement(Vml + "shadow", new XAttribute("on", "t"), new XAttribute("color", "black"), new XAttribute("obscured", "t")),
            new XElement(Vml + "path", new XAttribute(Office + "connecttype", "none")),
            new XElement(Vml + "textbox", new XAttribute("style", "mso-direction-alt:auto"), new XElement("div", new XAttribute("style", "text-align:left"))),
            new XElement(Excel + "ClientData", new XAttribute("ObjectType", "Note"),
                new XElement(Excel + "MoveWithCells"), new XElement(Excel + "SizeWithCells"),
                new XElement(Excel + "Anchor", anchor), new XElement(Excel + "AutoFill", "False"),
                new XElement(Excel + "Row", row), new XElement(Excel + "Column", column),
                visible ? new XElement(Excel + "Visible") : null));
    }

    // The idmap blocks VML drawings of the file use: shape ids must differ across the workbook.
    private static HashSet<int> UsedIdMaps(PackageEdit edit)
    {
        var used = new HashSet<int>();
        foreach (var part in edit.Package.Parts.Where(p => p.EndsWith(".vml", StringComparison.OrdinalIgnoreCase)))
        {
            var text = edit.Package.ReadText(part);
            foreach (Match match in IdMap().Matches(text))
            {
                foreach (var value in match.Groups[1].Value.Split(','))
                {
                    if (int.TryParse(value.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out var map))
                        used.Add(map);
                }
            }
        }

        return used;
    }

    [GeneratedRegex("<o:idmap[^>]*data=\"([0-9, ]+)\"")]
    private static partial Regex IdMap();
}
