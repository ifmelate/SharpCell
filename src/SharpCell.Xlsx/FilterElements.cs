using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using SharpCell.Evaluation;

namespace SharpCell.Xlsx;

/// <summary>
/// A sheet's <c>autoFilter</c> element and the hidden name <c>_xlnm._FilterDatabase</c> Excel keeps
/// its range in. Each filterColumn's colId counts from the range's first column.
/// </summary>
internal static class FilterElements
{
    public const string DatabaseName = "_xlnm._FilterDatabase";

    // The children of filterColumn that hold criteria; a filterColumn without one only hides its button.
    private static readonly HashSet<string> Criteria = new(StringComparer.Ordinal)
    {
        "filters", "customFilters", "top10", "dynamicFilter", "colorFilter", "iconFilter",
    };

    private static readonly string[] Operators = ["equal", "notEqual", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual"];

    private static readonly string[] Dynamics =
    [
        "aboveAverage", "belowAverage", "yesterday", "today", "tomorrow", "lastWeek", "thisWeek", "nextWeek",
        "lastMonth", "thisMonth", "nextMonth", "lastQuarter", "thisQuarter", "nextQuarter", "lastYear", "thisYear", "nextYear",
        "yearToDate", "Q1", "Q2", "Q3", "Q4", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12",
    ];

    private static readonly string[] Groupings = ["year", "month", "day", "hour", "minute", "second"];

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

    /// <summary>Reads a sheet's filter. One with a range SharpCell cannot read is left to the file.</summary>
    public static void Read(XElement element, Worksheet sheet)
    {
        var columns = element.Elements().Where(e => e.Name.LocalName == "filterColumn").ToList();
        if (columns.Exists(column => column.Elements().Any(e => Criteria.Contains(e.Name.LocalName))))
            sheet.FilterMode = true;
        if ((string?)element.Attribute("ref") is not { } reference
            || !ReferenceSyntax.TryParseA1Area(reference.Replace("$", "", StringComparison.Ordinal), new CellAddress(1, 1), out var parsed))
            return;
        var area = Area.Resolve(parsed, new CellAddress(1, 1));
        if (area.Rows < 2)
            return;
        var filter = sheet.LoadAutoFilter(area);
        foreach (var column in columns)
        {
            if (!int.TryParse((string?)column.Attribute("colId"), NumberStyles.None, CultureInfo.InvariantCulture, out var id) || id >= area.Columns)
                continue;
            filter.Load(area.FirstColumn + id, Column(column));
        }
    }

    // The criteria of a filterColumn.
    private static ColumnFilter Column(XElement column)
    {
        var criteria = column.Elements().FirstOrDefault(e => Criteria.Contains(e.Name.LocalName));
        switch (criteria?.Name.LocalName)
        {
            case "filters":
            {
                var values = criteria.Elements().Where(e => e.Name.LocalName == "filter").Select(e => (string?)e.Attribute("val") ?? "").ToArray();
                var dates = criteria.Elements().Where(e => e.Name.LocalName == "dateGroupItem").Select(Date).OfType<FilterDate>().ToArray();
                return (dates.Length > 0 ? ColumnFilter.OneOf(dates) with { Values = values } : ColumnFilter.OneOf(values))
                    with { Blanks = Flag(criteria, "blank", false) };
            }

            case "customFilters":
            {
                var conditions = criteria.Elements().Where(e => e.Name.LocalName == "customFilter")
                    .Select(e => new FilterCondition((FilterOperator)Math.Max(0, Array.IndexOf(Operators, (string?)e.Attribute("operator") ?? "equal")), (string?)e.Attribute("val") ?? ""))
                    .Take(2).ToList();
                if (conditions.Count == 0)
                    break;
                return conditions.Count == 1
                    ? ColumnFilter.Custom(conditions[0].Operator, conditions[0].Value)
                    : ColumnFilter.Custom(conditions[0], conditions[1], Flag(criteria, "and", false));
            }

            case "top10":
            {
                var percent = Flag(criteria, "percent", false);
                if (!double.TryParse((string?)criteria.Attribute("val"), NumberStyles.Float, CultureInfo.InvariantCulture, out var count)
                    || !(count >= 1 && count <= (percent ? 100 : 500)))
                    break;
                return Flag(criteria, "top", true) ? ColumnFilter.Top(count, percent) : ColumnFilter.Bottom(count, percent);
            }

            case "dynamicFilter":
            {
                var index = Array.IndexOf(Dynamics, (string?)criteria.Attribute("type"));
                if (index < 0)
                    break;
                return ColumnFilter.Of((DynamicFilter)index);
            }
        }

        // By colour or icon: kept, and not applied. Without criteria SharpCell can read (only hiding
        // the button): kept, and filtering nothing.
        return criteria?.Name.LocalName is "colorFilter" or "iconFilter"
            ? ColumnFilter.ByColor(new XElement(column))
            : ColumnFilter.OneOf(Array.Empty<string>()) with { Source = new XElement(column) };
    }

    private static FilterDate? Date(XElement item)
    {
        var grouping = Array.IndexOf(Groupings, (string?)item.Attribute("dateTimeGrouping"));
        if (grouping < 0 || Int(item, "year") is not { } year)
            return null;
        return new FilterDate((DateGrouping)grouping, year, Int(item, "month") ?? 1, Int(item, "day") ?? 1, Int(item, "hour") ?? 0,
            Int(item, "minute") ?? 0, Int(item, "second") ?? 0);
    }

    private static int? Int(XElement element, string name) =>
        int.TryParse((string?)element.Attribute(name), NumberStyles.None, CultureInfo.InvariantCulture, out var value) ? value : null;

    private static bool Flag(XElement element, string name, bool absent) => (string?)element.Attribute(name) switch
    {
        null => absent,
        var text => text is "1" or "true",
    };

    /// <summary>
    /// The sheet's autoFilter element: the file's (keeping its sort state and extensions) with the
    /// model's range and criteria, or a new one; null when the sheet has no filter.
    /// </summary>
    public static XElement? Element(XNamespace ns, Worksheet sheet, XElement? source)
    {
        if (sheet.AutoFilter is not { } filter)
            return null;
        var area = filter.Area;
        var element = source ?? new XElement(ns + "autoFilter");
        element.SetAttributeValue("ref", WorksheetParts.Format(area));
        element.Elements().Where(e => e.Name.LocalName == "filterColumn").Remove();
        foreach (var (column, criteria) in filter.Columns.Reverse())
        {
            var id = (column - area.FirstColumn).ToString(CultureInfo.InvariantCulture);
            if (criteria.Source is XElement kept)
            {
                var copy = new XElement(kept);
                copy.Name = ns + "filterColumn";
                copy.SetAttributeValue("colId", id);
                element.AddFirst(copy);
                continue;
            }

            if (!criteria.HasCriteria)
                continue;
            element.AddFirst(new XElement(ns + "filterColumn", new XAttribute("colId", id),
                Criterion(ns, criteria, filter.Limits.TryGetValue(column, out var limit) ? limit : null)));
        }

        return element;
    }

    private static XElement Criterion(XNamespace ns, ColumnFilter filter, (double Value, double? Max)? limit)
    {
        switch (filter.Type)
        {
            case ColumnFilterType.Values:
                return new XElement(ns + "filters", filter.Blanks ? new XAttribute("blank", "1") : null,
                    filter.Values.Select(value => new XElement(ns + "filter", new XAttribute("val", value))),
                    filter.Dates.Select(date => DateItem(ns, date)));
            case ColumnFilterType.Custom:
                return new XElement(ns + "customFilters", filter.MatchAll && filter.Conditions.Count > 1 ? new XAttribute("and", "1") : null,
                    filter.Conditions.Select(condition => new XElement(ns + "customFilter",
                        condition.Operator == FilterOperator.Equal ? null : new XAttribute("operator", Operators[(int)condition.Operator]),
                        new XAttribute("val", condition.Value))));
            case ColumnFilterType.Top or ColumnFilterType.Bottom:
                return new XElement(ns + "top10", filter.Type == ColumnFilterType.Bottom ? new XAttribute("top", "0") : null,
                    filter.Percent ? new XAttribute("percent", "1") : null, new XAttribute("val", Number(filter.Count)),
                    limit is { } top ? new XAttribute("filterVal", Number(top.Value)) : null);
            default:
                return new XElement(ns + "dynamicFilter", new XAttribute("type", Dynamics[(int)filter.Dynamic]),
                    limit is { } dynamic ? new XAttribute("val", Number(dynamic.Value)) : null,
                    limit is { Max: { } max } ? new XAttribute("maxVal", Number(max)) : null);
        }
    }

    private static XElement DateItem(XNamespace ns, FilterDate date)
    {
        var element = new XElement(ns + "dateGroupItem", new XAttribute("year", date.Year));
        if (date.Grouping >= DateGrouping.Month)
            element.Add(new XAttribute("month", date.Month));
        if (date.Grouping >= DateGrouping.Day)
            element.Add(new XAttribute("day", date.Day));
        if (date.Grouping >= DateGrouping.Hour)
            element.Add(new XAttribute("hour", date.Hour));
        if (date.Grouping >= DateGrouping.Minute)
            element.Add(new XAttribute("minute", date.Minute));
        if (date.Grouping >= DateGrouping.Second)
            element.Add(new XAttribute("second", date.Second));
        element.Add(new XAttribute("dateTimeGrouping", Groupings[(int)date.Grouping]));
        return element;
    }

    private static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);

    /// <summary>The formula of the sheet's filter database name: its range, absolute; null without a filter.</summary>
    public static string? DatabaseFormula(Worksheet sheet) =>
        sheet.AutoFilter is { } filter ? PrintNames.Reference(sheet, filter.Area) : null;
}
