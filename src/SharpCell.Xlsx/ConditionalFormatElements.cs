using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Xlsx;

/// <summary>
/// A sheet's conditional formats in its &lt;conditionalFormatting&gt; elements: each with its ranges in
/// sqref and its rules (&lt;cfRule&gt;), whose formulas are relative to the first cell of sqref and whose
/// look is a differential format (dxf) of the styles part. Rules in the extension list
/// (x14:conditionalFormattings) are not read; the file keeps them.
/// </summary>
internal static class ConditionalFormatElements
{
    /// <summary>What the file had for a rule: its element, written back with the model's settings, and its dxf.</summary>
    internal sealed record Source(XElement Rule, int? DxfId);

    /// <summary>
    /// The rules of one &lt;conditionalFormatting&gt; element with their priorities, to be added in
    /// priority order once the sheet is read. A rule Excel would not read is skipped.
    /// </summary>
    public static IEnumerable<(int Priority, Action<Worksheet> Add)> Read(XElement element, IReadOnlyList<ConditionalStyle> dxfs)
    {
        if (Areas((string?)element.Attribute("sqref")) is not { } areas)
            yield break;
        var origin = new CellAddress(areas[0].FirstRow, areas[0].FirstColumn);
        var ns = element.Name.Namespace;
        foreach (var rule in element.Elements(ns + "cfRule"))
        {
            var dxfId = Int(rule, "dxfId");
            var style = dxfId is { } id && id >= 0 && id < dxfs.Count ? dxfs[id] : null;
            if (Format(rule, style, origin) is not { } read)
                continue;
            var priority = Int(rule, "priority") ?? int.MaxValue;
            var source = new Source(new XElement(rule), dxfId);
            yield return (priority, sheet => sheet.AddConditionalRule(areas, read.Format, read.Condition, source));
        }
    }

    private static (ConditionalFormat Format, string? Condition)? Format(XElement rule, ConditionalStyle? style, CellAddress origin)
    {
        var ns = rule.Name.Namespace;
        var formulas = rule.Elements(ns + "formula").Select(f => Formula(f, origin)).ToList();
        var first = formulas.ElementAtOrDefault(0);
        var second = formulas.ElementAtOrDefault(1);
        var stop = CellMetadata.IsTrue((string?)rule.Attribute("stopIfTrue"));
        style ??= new ConditionalStyle();
        ConditionalFormat Make(ConditionalFormatType type, ConditionalOperator comparison = ConditionalOperator.Between, string? f1 = null, string? f2 = null,
            string? text = null, int rank = 0, bool percent = false, bool orEqual = false, int deviations = 0, TimePeriod period = TimePeriod.Yesterday) =>
            ConditionalFormat.Read(type, style, comparison, f1, f2, text, rank, percent, orEqual, deviations, period, stop);

        switch ((string?)rule.Attribute("type"))
        {
            case "cellIs" when first is not null && OperatorOf((string?)rule.Attribute("operator") ?? "between") is { } comparison:
                var between = comparison is ConditionalOperator.Between or ConditionalOperator.NotBetween;
                if (between && second is null)
                    return null;
                return (Make(ConditionalFormatType.CellValue, comparison, first, between ? second : null), null);
            case "expression" when first is not null:
                return (Make(ConditionalFormatType.Expression, f1: first), null);
            case "top10":
                var rank = Int(rule, "rank") ?? 10;
                var percent = CellMetadata.IsTrue((string?)rule.Attribute("percent"));
                if (rank < 1 || rank > (percent ? 100 : 1000))
                    return null;
                var bottom = CellMetadata.IsTrue((string?)rule.Attribute("bottom"));
                return (Make(bottom ? ConditionalFormatType.Bottom : ConditionalFormatType.Top, rank: rank, percent: percent), null);
            case "aboveAverage":
                var above = (string?)rule.Attribute("aboveAverage") is not ("0" or "false");
                var deviations = Int(rule, "stdDev") ?? 0;
                if (deviations is < 0 or > 3)
                    return null;
                return (Make(above ? ConditionalFormatType.AboveAverage : ConditionalFormatType.BelowAverage,
                    orEqual: CellMetadata.IsTrue((string?)rule.Attribute("equalAverage")), deviations: deviations), null);
            case "duplicateValues":
                return (Make(ConditionalFormatType.DuplicateValues), null);
            case "uniqueValues":
                return (Make(ConditionalFormatType.UniqueValues), null);
            case "containsText" or "notContainsText" or "beginsWith" or "endsWith" when (string?)rule.Attribute("text") is { Length: > 0 } text:
                var type = (string?)rule.Attribute("type") switch
                {
                    "containsText" => ConditionalFormatType.ContainsText,
                    "notContainsText" => ConditionalFormatType.NotContainsText,
                    "beginsWith" => ConditionalFormatType.BeginsWith,
                    _ => ConditionalFormatType.EndsWith,
                };
                return (Make(type, text: text), first);
            case "containsBlanks":
                return (Make(ConditionalFormatType.Blanks), first);
            case "notContainsBlanks":
                return (Make(ConditionalFormatType.NoBlanks), first);
            case "containsErrors":
                return (Make(ConditionalFormatType.Errors), first);
            case "notContainsErrors":
                return (Make(ConditionalFormatType.NoErrors), first);
            case "timePeriod" when PeriodOf((string?)rule.Attribute("timePeriod")) is { } period:
                return (Make(ConditionalFormatType.TimePeriod, period: period), first);
            case "colorScale":
                return (Make(ConditionalFormatType.ColorScale) with { Style = null }, null);
            case "dataBar":
                return (Make(ConditionalFormatType.DataBar) with { Style = null }, null);
            case "iconSet":
                return (Make(ConditionalFormatType.IconSet) with { Style = null }, null);
            default:
                return null;
        }
    }

    /// <summary>The sheet's &lt;conditionalFormatting&gt; elements, rules of the same ranges together, in priority order.</summary>
    public static List<XElement> Elements(XNamespace ns, Worksheet sheet, StylesWriter styles)
    {
        var elements = new List<XElement>();
        XElement? current = null;
        IReadOnlyList<Area>? currentAreas = null;
        var priority = 1;
        foreach (var rule in sheet.ConditionalRules)
        {
            if (current is null || !rule.Areas.SequenceEqual(currentAreas!))
            {
                current = new XElement(ns + "conditionalFormatting", new XAttribute("sqref", string.Join(' ', rule.Areas.Select(WorksheetParts.Format))));
                currentAreas = rule.Areas;
                elements.Add(current);
            }

            current.Add(Rule(ns, rule, priority++, sheet, styles));
        }

        return elements;
    }

    private static XElement Rule(XNamespace ns, ConditionalRule rule, int priority, Worksheet sheet, StylesWriter styles)
    {
        var format = rule.Format;
        var source = rule.Source as Source;
        var element = source is not null ? new XElement(source.Rule) : new XElement(ns + "cfRule");
        foreach (var name in new[] { "type", "dxfId", "priority", "stopIfTrue", "operator", "text", "timePeriod", "rank", "percent", "bottom", "aboveAverage", "equalAverage", "stdDev" })
            element.Attribute(name)?.Remove();
        element.Elements(ns + "formula").Remove();

        var attributes = new List<XAttribute> { new("type", TypeText(format)) };
        if (format.Style is { } style)
        {
            var dxfId = source?.DxfId is { } id && styles.DxfAt(id) == style ? id : styles.DxfFor(style);
            attributes.Add(new XAttribute("dxfId", dxfId));
        }

        attributes.Add(new XAttribute("priority", priority));
        if (format.StopIfTrue)
            attributes.Add(new XAttribute("stopIfTrue", "1"));
        switch (format.Type)
        {
            case ConditionalFormatType.CellValue:
                attributes.Add(new XAttribute("operator", OperatorText(format.Operator)));
                break;
            case ConditionalFormatType.Top or ConditionalFormatType.Bottom:
                attributes.Add(new XAttribute("rank", format.Rank));
                if (format.Percent)
                    attributes.Add(new XAttribute("percent", "1"));
                if (format.Type == ConditionalFormatType.Bottom)
                    attributes.Add(new XAttribute("bottom", "1"));
                break;
            case ConditionalFormatType.AboveAverage or ConditionalFormatType.BelowAverage:
                if (format.Type == ConditionalFormatType.BelowAverage)
                    attributes.Add(new XAttribute("aboveAverage", "0"));
                if (format.OrEqual)
                    attributes.Add(new XAttribute("equalAverage", "1"));
                if (format.StandardDeviations > 0)
                    attributes.Add(new XAttribute("stdDev", format.StandardDeviations));
                break;
            case ConditionalFormatType.ContainsText or ConditionalFormatType.NotContainsText or ConditionalFormatType.BeginsWith or ConditionalFormatType.EndsWith:
                attributes.Add(new XAttribute("operator", format.Type switch
                {
                    ConditionalFormatType.ContainsText => "containsText",
                    ConditionalFormatType.NotContainsText => "notContains",
                    ConditionalFormatType.BeginsWith => "beginsWith",
                    _ => "endsWith",
                }));
                attributes.Add(new XAttribute("text", format.Text!));
                break;
            case ConditionalFormatType.TimePeriod:
                attributes.Add(new XAttribute("timePeriod", PeriodText(format.Period)));
                break;
        }

        // The model's attributes first, in the schema's order, then whatever else the file's element had.
        var others = element.Attributes().ToList();
        element.RemoveAttributes();
        element.Add(attributes, others);
        element.AddFirst(new[] { rule.F1, rule.F2 }.Where(f => f is not null).Select(f => FormulaElement(ns, f!, rule.Origin, sheet)));
        return element;
    }

    // A formula SharpCell could not read is written as the file had it.
    private static XElement FormulaElement(XNamespace ns, FormulaNode formula, CellAddress origin, Worksheet sheet) =>
        new(ns + "formula", FormulaPrinter.PrintForFile(formula, origin, at => sheet.Workbook.TableAt(sheet, at.Row, at.Column)?.Name,
            function => FunctionValues.IsFunction(sheet.Workbook, sheet, function)));

    private static string? Formula(XElement element, CellAddress origin)
    {
        if (element.Value is not { Length: > 0 } text)
            return null;
        if (FormulaText.ReferencesOtherWorkbook(text))
            return "=" + text;
        try
        {
            return "=" + FormulaPrinter.Print(FormulaParser.Parse("=" + text, origin), origin);
        }
        catch (FormulaParseException)
        {
            return "=" + text;
        }
    }

    private static List<Area>? Areas(string? sqref)
    {
        var areas = new List<Area>();
        foreach (var part in (sqref ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (!ReferenceSyntax.TryParseA1Area(part.Replace("$", "", StringComparison.Ordinal), new CellAddress(1, 1), out var parsed))
                return null;
            areas.Add(Area.Resolve(parsed, new CellAddress(1, 1)));
        }

        return areas.Count > 0 ? areas : null;
    }

    private static int? Int(XElement element, string name) =>
        int.TryParse((string?)element.Attribute(name), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var value) ? value : null;

    private static string TypeText(ConditionalFormat format) => format.Type switch
    {
        ConditionalFormatType.CellValue => "cellIs",
        ConditionalFormatType.Expression => "expression",
        ConditionalFormatType.Top or ConditionalFormatType.Bottom => "top10",
        ConditionalFormatType.AboveAverage or ConditionalFormatType.BelowAverage => "aboveAverage",
        ConditionalFormatType.DuplicateValues => "duplicateValues",
        ConditionalFormatType.UniqueValues => "uniqueValues",
        ConditionalFormatType.ContainsText => "containsText",
        ConditionalFormatType.NotContainsText => "notContainsText",
        ConditionalFormatType.BeginsWith => "beginsWith",
        ConditionalFormatType.EndsWith => "endsWith",
        ConditionalFormatType.Blanks => "containsBlanks",
        ConditionalFormatType.NoBlanks => "notContainsBlanks",
        ConditionalFormatType.Errors => "containsErrors",
        ConditionalFormatType.NoErrors => "notContainsErrors",
        ConditionalFormatType.TimePeriod => "timePeriod",
        ConditionalFormatType.ColorScale => "colorScale",
        ConditionalFormatType.DataBar => "dataBar",
        _ => "iconSet",
    };

    private static readonly string[] Operators = ["between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"];

    private static ConditionalOperator? OperatorOf(string text) => Array.IndexOf(Operators, text) is var i and >= 0 ? (ConditionalOperator)i : null;

    private static string OperatorText(ConditionalOperator comparison) => Operators[(int)comparison];

    private static readonly string[] Periods = ["yesterday", "today", "tomorrow", "last7Days", "lastWeek", "thisWeek", "nextWeek", "lastMonth", "thisMonth", "nextMonth"];

    private static TimePeriod? PeriodOf(string? text) => text is not null && Array.IndexOf(Periods, text) is var i and >= 0 ? (TimePeriod)i : null;

    private static string PeriodText(TimePeriod period) => Periods[(int)period];
}
