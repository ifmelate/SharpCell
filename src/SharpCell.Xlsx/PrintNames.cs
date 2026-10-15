using System;
using System.Collections.Generic;
using System.Linq;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Xlsx;

/// <summary>
/// The built-in names Excel keeps a sheet's print area and print titles in: <c>_xlnm.Print_Area</c>
/// and <c>_xlnm.Print_Titles</c>, scoped to the sheet, with absolute references to it.
/// </summary>
internal static class PrintNames
{
    public const string Area = "_xlnm.Print_Area";
    public const string Titles = "_xlnm.Print_Titles";

    public static bool IsPrintName(string name) => IsArea(name) || name.Equals(Titles, StringComparison.OrdinalIgnoreCase);

    public static bool IsArea(string name) => name.Equals(Area, StringComparison.OrdinalIgnoreCase);

    /// <summary>The ranges of a name's formula, such as <c>'My Sheet'!$A$1:$F$40,'My Sheet'!$H$1:$J$5</c>; null when it is anything else.</summary>
    public static List<Area>? Parse(string formula)
    {
        var areas = new List<Area>();
        foreach (var part in Split(formula.TrimStart('=')))
        {
            var reference = WithoutSheet(part.Trim());
            if (reference is null
                || !ReferenceSyntax.TryParseA1Area(reference.Replace("$", "", StringComparison.Ordinal).AsSpan(), new CellAddress(1, 1), out var parsed))
                return null;
            areas.Add(SharpCell.Evaluation.Area.Resolve(parsed, new CellAddress(1, 1)));
        }

        return areas.Count == 0 ? null : areas;
    }

    /// <summary>The formula of the sheet's print area; null when it has none.</summary>
    public static string? AreaFormula(Worksheet sheet) =>
        sheet.PrintAreas.Count == 0 ? null : string.Join(",", sheet.PrintAreas.Select(area => Reference(sheet, area)));

    /// <summary>The formula of the sheet's print titles, columns first as Excel writes them; null when it has none.</summary>
    public static string? TitlesFormula(Worksheet sheet)
    {
        var parts = new List<string>();
        if (sheet.TitleColumns is { } columns)
            parts.Add(Reference(sheet, new Area(1, columns.First, CellAddress.MaxRow, columns.Last)));
        if (sheet.TitleRows is { } rows)
            parts.Add(Reference(sheet, new Area(rows.First, 1, rows.Last, CellAddress.MaxColumn)));
        return parts.Count == 0 ? null : string.Join(",", parts);
    }

    private static string Reference(Worksheet sheet, Area area)
    {
        var prefix = FormulaPrinter.QuoteSheetName(sheet.Name) + "!";
        if (area.FirstColumn == 1 && area.LastColumn == CellAddress.MaxColumn)
            return $"{prefix}${area.FirstRow}:${area.LastRow}";
        if (area.FirstRow == 1 && area.LastRow == CellAddress.MaxRow)
            return $"{prefix}${ColumnNames.ToName(area.FirstColumn)}:${ColumnNames.ToName(area.LastColumn)}";
        var first = $"${ColumnNames.ToName(area.FirstColumn)}${area.FirstRow}";
        return area.IsSingleCell ? prefix + first : $"{prefix}{first}:${ColumnNames.ToName(area.LastColumn)}${area.LastRow}";
    }

    // The parts of a list separated by commas outside quoted sheet names.
    private static IEnumerable<string> Split(string text)
    {
        var start = 0;
        var quoted = false;
        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] == '\'')
                quoted = !quoted;
            else if (text[i] == ',' && !quoted)
            {
                yield return text[start..i];
                start = i + 1;
            }
        }

        yield return text[start..];
    }

    // The reference after its sheet name, quoted (with '' for a quote) or not; null without a sheet name.
    private static string? WithoutSheet(string part)
    {
        if (part.StartsWith('\''))
        {
            var i = 1;
            while (i < part.Length)
            {
                if (part[i] == '\'' && i + 1 < part.Length && part[i + 1] == '\'')
                    i += 2;
                else if (part[i] == '\'')
                    break;
                else
                    i++;
            }

            return i + 1 < part.Length && part[i + 1] == '!' ? part[(i + 2)..] : null;
        }

        var bang = part.LastIndexOf('!');
        return bang < 0 ? null : part[(bang + 1)..];
    }
}
