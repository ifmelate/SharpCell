using System;
using System.Globalization;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace SharpCell.Xlsx;

/// <summary>
/// Parts of the file that point at cells of a sheet whose rows or columns moved: drawings
/// (pictures, shapes, charts) by their anchors, comments and their boxes, threaded comments,
/// chart series formulas, and pivot table sources and locations. Each follows the cells as Excel
/// moves them; a comment on a deleted cell goes, and an object anchored in deleted rows shrinks to
/// where they were.
/// </summary>
internal static class RelatedPartsShift
{
    private const string ChartType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    private const string ChartExType = "application/vnd.ms-office.chartex+xml";
    private const string PivotCacheType = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";

    public static void Apply(PackageEdit edit, WritePlan plan)
    {
        var map = plan.Edits;
        if (!map.Any)
            return;
        var package = edit.Package;
        foreach (var (sheet, source) in plan.Source.LiveSheets())
        {
            if (!map.HasEdits(sheet))
                continue;
            foreach (var relationship in package.ReadRelationships(source.Part!).Values)
            {
                if (!package.Exists(relationship.Target))
                    continue;
                if (relationship.Is("drawing"))
                    Update(edit, relationship.Target, root => Drawing(root, map, sheet));
                else if (relationship.Is("comments"))
                    Update(edit, relationship.Target, root => Comments(root, "comment", map, sheet));
                else if (relationship.Is("threadedComment"))
                    Update(edit, relationship.Target, root => Comments(root, "threadedComment", map, sheet));
                else if (relationship.Is("vmlDrawing"))
                    Update(edit, relationship.Target, root => Notes(root, map, sheet));
                else if (relationship.Is("pivotTable"))
                    Update(edit, relationship.Target, root => PivotLocation(root, map, sheet));
            }
        }

        var types = edit.ReadXml("[Content_Types].xml")?.Root;
        foreach (var part in types?.Elements().Where(e => e.Name.LocalName == "Override")
                     .Select(e => (Part: ((string?)e.Attribute("PartName"))?.TrimStart('/'), Type: (string?)e.Attribute("ContentType")))
                     .Where(p => p.Part is not null && package.Exists(p.Part)) ?? [])
        {
            if (part.Type is ChartType or ChartExType)
                Update(edit, part.Part!, root => ChartFormulas(root, map));
            else if (part.Type == PivotCacheType)
                Update(edit, part.Part!, root => PivotSource(root, map, plan.Workbook));
        }
    }

    // Loads the part, lets the change work on it, and keeps it when it changed. A part that is not
    // well-formed XML (some writers' VML) is left as it is.
    private static void Update(PackageEdit edit, string part, Action<XElement> change)
    {
        XDocument? document;
        try
        {
            document = edit.ReadXml(part);
        }
        catch (XmlException)
        {
            return;
        }

        if (document?.Root is not { } root)
            return;
        var before = new XElement(root);
        change(root);
        if (!XNode.DeepEquals(before, root))
            edit.Parts[part] = XmlParts.Save(document);
    }

    // Anchors count rows and columns from 0. A two-cell anchor moves and sizes with its cells, a
    // one-cell one moves, an absolute one stays.
    private static void Drawing(XElement root, SheetEditMap map, Worksheet sheet)
    {
        foreach (var anchor in root.Elements().Where(e => e.Name.LocalName is "twoCellAnchor" or "oneCellAnchor"))
        {
            var editAs = (string?)anchor.Attribute("editAs");
            if (editAs == "absolute")
                continue;
            var from = anchor.Elements().FirstOrDefault(e => e.Name.LocalName == "from");
            var to = anchor.Elements().FirstOrDefault(e => e.Name.LocalName == "to");
            var (fromRow, fromColumn) = (Index(from, "row"), Index(from, "col"));
            if (fromRow is null || fromColumn is null)
                continue;
            var newRow = Clamped(map, sheet, fromRow.Value + 1, rows: true) - 1;
            var newColumn = Clamped(map, sheet, fromColumn.Value + 1, rows: false) - 1;
            if (Index(to, "row") is { } toRow && Index(to, "col") is { } toColumn)
            {
                if (editAs == "oneCell")
                {
                    SetIndex(to!, "row", toRow + newRow - fromRow.Value);
                    SetIndex(to!, "col", toColumn + newColumn - fromColumn.Value);
                }
                else
                {
                    SetIndex(to!, "row", Math.Max(newRow, Clamped(map, sheet, toRow + 1, rows: true) - 1));
                    SetIndex(to!, "col", Math.Max(newColumn, Clamped(map, sheet, toColumn + 1, rows: false) - 1));
                }
            }

            SetIndex(from!, "row", newRow);
            SetIndex(from!, "col", newColumn);
        }
    }

    // Where a row or column goes; a deleted one is where the deletion was.
    private static int Clamped(SheetEditMap map, Worksheet sheet, int index, bool rows)
    {
        var current = index;
        foreach (var edit in sheet.Edits)
        {
            // Shapes stay when only cells beside them shift.
            if (edit.Rows != rows || !edit.Whole)
                continue;
            var max = rows ? CellAddress.MaxRow : CellAddress.MaxColumn;
            current = edit.Map(current, max) ?? (edit.Delete ? Math.Min(edit.At, max) : max);
        }

        return current;
    }

    private static int? Index(XElement? marker, string name) =>
        marker?.Elements().FirstOrDefault(e => e.Name.LocalName == name) is { } element
        && int.TryParse(element.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var index) ? index : null;

    private static void SetIndex(XElement marker, string name, int value) =>
        marker.Elements().First(e => e.Name.LocalName == name).Value = value.ToString(CultureInfo.InvariantCulture);

    // Comments (and threaded comments) name their cell in ref; those on deleted cells go.
    private static void Comments(XElement root, string name, SheetEditMap map, Worksheet sheet)
    {
        foreach (var comment in root.Descendants().Where(e => e.Name.LocalName == name).ToList())
        {
            if ((string?)comment.Attribute("ref") is not { } reference || !CellAddress.TryParse(reference, out var cell))
                continue;
            if (map.Map(sheet, cell) is { } moved)
                comment.SetAttributeValue("ref", moved.ToString());
            else
                comment.Remove();
        }
    }

    // The boxes of legacy comments: x:Row and x:Column (from 0) name the cell, x:Anchor places the box.
    private static void Notes(XElement root, SheetEditMap map, Worksheet sheet)
    {
        foreach (var data in root.Descendants().Where(e => e.Name.LocalName == "ClientData").ToList())
        {
            var row = data.Elements().FirstOrDefault(e => e.Name.LocalName == "Row");
            var column = data.Elements().FirstOrDefault(e => e.Name.LocalName == "Column");
            if (!int.TryParse(row?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var r)
                || !int.TryParse(column?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var c))
                continue;
            if (map.Map(sheet, new CellAddress(r + 1, c + 1)) is not { } moved)
            {
                // The shape holding the note goes with it.
                (data.Parent ?? data).Remove();
                continue;
            }

            var (rowShift, columnShift) = (moved.Row - 1 - r, moved.Column - 1 - c);
            row!.Value = (moved.Row - 1).ToString(CultureInfo.InvariantCulture);
            column!.Value = (moved.Column - 1).ToString(CultureInfo.InvariantCulture);
            if (data.Elements().FirstOrDefault(e => e.Name.LocalName == "Anchor") is { } anchor)
            {
                var parts = anchor.Value.Split(',').Select(p => p.Trim()).ToArray();
                if (parts.Length == 8 && parts.All(p => int.TryParse(p, NumberStyles.None, CultureInfo.InvariantCulture, out _)))
                {
                    int Part(int i) => int.Parse(parts[i], CultureInfo.InvariantCulture);
                    var values = new[]
                    {
                        Part(0) + columnShift, Part(1), Part(2) + rowShift, Part(3), Part(4) + columnShift, Part(5), Part(6) + rowShift, Part(7),
                    };
                    anchor.Value = string.Join(", ", values.Select(v => Math.Max(0, v).ToString(CultureInfo.InvariantCulture)));
                }
            }
        }
    }

    private static void ChartFormulas(XElement root, SheetEditMap map)
    {
        foreach (var formula in root.Descendants().Where(e => e.Name.LocalName == "f" && !e.HasElements))
            formula.Value = map.MapFormula(formula.Value, null, new CellAddress(1, 1));
    }

    private static void PivotSource(XElement root, SheetEditMap map, Workbook workbook)
    {
        foreach (var source in root.Descendants().Where(e => e.Name.LocalName == "worksheetSource"))
        {
            // A source on a removed sheet stays; Excel reports it when the pivot table is refreshed.
            if ((string?)source.Attribute("sheet") is not { } name || map.SheetName(name) is not { } current)
                continue;
            if (current != name)
                source.SetAttributeValue("sheet", current);
            if ((string?)source.Attribute("ref") is { } reference
                && workbook.TryGetSheet(current, out var sheet) && map.HasEdits(sheet!)
                && map.MapSqref(sheet!, reference) is { } moved)
                source.SetAttributeValue("ref", moved);
        }
    }

    // A pivot table moves with its rows, but Excel does not let rows or columns be inserted into it
    // or deleted from it: its definition would no longer fit its place.
    private static void PivotLocation(XElement root, SheetEditMap map, Worksheet sheet)
    {
        if (root.Elements().FirstOrDefault(e => e.Name.LocalName == "location") is not { } location
            || (string?)location.Attribute("ref") is not { } reference
            || !ReferenceSyntax.TryParseA1Area(reference, new CellAddress(1, 1), out var parsed))
            return;
        var area = SharpCell.Evaluation.Area.Resolve(parsed, new CellAddress(1, 1));
        if (map.Map(sheet, area) is not { } moved || moved.Rows != area.Rows || moved.Columns != area.Columns)
            throw new NotSupportedException($"Rows or columns were inserted into or deleted from pivot table '{(string?)root.Attribute("name")}' on sheet '{sheet.Name}'; Excel does not allow that.");
        location.SetAttributeValue("ref", WorksheetParts.Format(moved));
    }
}
