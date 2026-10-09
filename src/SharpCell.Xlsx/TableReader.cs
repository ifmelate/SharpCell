using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using SharpCell.Evaluation;

namespace SharpCell.Xlsx;

/// <summary>A table part: what SharpCell needs from <c>xl/tables/tableN.xml</c>.</summary>
internal sealed record TablePart(string Name, Area Area, bool HasHeaderRow, bool HasTotalsRow, IReadOnlyList<string> Columns, bool Filtered);

internal static class TableReader
{
    public static TablePart Read(Package package, string part)
    {
        string? name = null, range = null;
        var header = true;
        var totals = false;
        var filtered = false;
        var columns = new List<string>();
        using var reader = package.OpenXml(part);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element)
                continue;
            switch (reader.LocalName)
            {
                case "table" when reader.Depth == 0:
                    // Formulas use displayName; name is a fallback some writers fill alone.
                    name = reader.GetAttribute("displayName") ?? reader.GetAttribute("name");
                    range = reader.GetAttribute("ref");
                    header = RowCount(reader.GetAttribute("headerRowCount"), 1, part) == 1;
                    totals = RowCount(reader.GetAttribute("totalsRowCount"), 0, part) == 1;
                    break;
                case "autoFilter" when reader.Depth == 1:
                    filtered |= AutoFilter.HasCriteria(reader);
                    break;
                case "tableColumn":
                    columns.Add(XmlText.Decode(reader.GetAttribute("name")
                        ?? throw new InvalidDataException($"A column of table part '{part}' has no name.")));
                    break;
            }
        }

        if (name is null || range is null)
            throw new InvalidDataException($"Table part '{part}' has no name or range.");
        if (!ReferenceSyntax.TryParseA1Area(range, new CellAddress(1, 1), out var area) || area.Kind != AreaKind.Range)
            throw new InvalidDataException($"Table range '{range}' in '{part}' is not valid.");
        return new TablePart(name, Area.Resolve(area, new CellAddress(1, 1)), header, totals, columns, filtered);
    }

    private static int RowCount(string? text, int fallback, string part) => text switch
    {
        null => fallback,
        "0" => 0,
        "1" => 1,
        _ => throw new InvalidDataException($"Table part '{part}' has a header or totals row count of '{text}'; only 0 and 1 are supported."),
    };
}

internal static class AutoFilter
{
    // The children of filterColumn that hold criteria; a filterColumn without one only hides its button.
    private static readonly HashSet<string> Criteria = new(StringComparer.Ordinal)
    {
        "filters", "customFilters", "top10", "dynamicFilter", "colorFilter", "iconFilter",
    };

    /// <summary>Whether the autoFilter element the reader is on has a column with criteria. Leaves the reader on its end.</summary>
    public static bool HasCriteria(XmlReader reader)
    {
        if (reader.IsEmptyElement)
            return false;
        using var subtree = reader.ReadSubtree();
        var found = false;
        while (subtree.Read())
        {
            // autoFilter is depth 0 here, filterColumn 1, criteria 2.
            if (subtree.NodeType == XmlNodeType.Element && subtree.Depth == 2 && Criteria.Contains(subtree.LocalName))
                found = true;
        }

        return found;
    }
}
