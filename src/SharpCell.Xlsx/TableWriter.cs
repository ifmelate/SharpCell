using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using SharpCell.Evaluation;

namespace SharpCell.Xlsx;

/// <summary>Table parts for tables added since the workbook was read, in the shape Excel writes.</summary>
internal static class TableWriter
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

    /// <summary>The ids of the package's tables: a new table needs one no other table has.</summary>
    public static HashSet<int> UsedIds(PackageEdit edit)
    {
        var ids = new HashSet<int>();
        var types = edit.ReadXml("[Content_Types].xml");
        foreach (var part in types?.Root?.Elements().Where(e => e.Name.LocalName == "Override" && (string?)e.Attribute("ContentType") == ContentType)
                     .Select(e => ((string?)e.Attribute("PartName"))?.TrimStart('/')) ?? [])
        {
            if (part is not null && edit.ReadXml(part)?.Root?.Attribute("id") is { } id && int.TryParse(id.Value, out var value))
                ids.Add(value);
        }

        return ids;
    }

    public static XDocument Write(Table table, int id)
    {
        var area = table.Area;
        var root = new XElement(Main + "table", new XAttribute("id", id), new XAttribute("name", table.Name),
            new XAttribute("displayName", table.Name), new XAttribute("ref", WorksheetParts.Format(area)));
        if (!table.HasHeaderRow)
            root.Add(new XAttribute("headerRowCount", 0));
        if (table.HasTotalsRow)
            root.Add(new XAttribute("totalsRowCount", 1));
        else
            root.Add(new XAttribute("totalsRowShown", 0));
        if (table.HasHeaderRow)
        {
            var filtered = table.HasTotalsRow ? area with { LastRow = area.LastRow - 1 } : area;
            root.Add(new XElement(Main + "autoFilter", new XAttribute("ref", WorksheetParts.Format(filtered))));
        }

        root.Add(new XElement(Main + "tableColumns", new XAttribute("count", table.Columns.Count),
            table.Columns.Select((name, i) => new XElement(Main + "tableColumn", new XAttribute("id", i + 1), new XAttribute("name", XmlText.Encode(name))))));
        root.Add(new XElement(Main + "tableStyleInfo", new XAttribute("name", "TableStyleMedium2"), new XAttribute("showFirstColumn", 0),
            new XAttribute("showLastColumn", 0), new XAttribute("showRowStripes", 1), new XAttribute("showColumnStripes", 0)));
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
    }

    /// <summary>
    /// Header cells that must be written with their column's name: Excel repairs a file whose table
    /// header cells do not hold the column names as text.
    /// </summary>
    public static IEnumerable<(CellAddress Address, string Name)> HeaderCells(Table table)
    {
        if (!table.HasHeaderRow)
            yield break;
        var row = table.Area.FirstRow;
        for (var i = 0; i < table.Columns.Count; i++)
        {
            var column = table.Area.FirstColumn + i;
            var data = table.Worksheet.Store.Get(row, column);
            if (data?.Formula is not null || data?.Value is not { Kind: CellValueKind.Text } value || value.AsText() != table.Columns[i])
                yield return (new CellAddress(row, column), table.Columns[i]);
        }
    }
}
