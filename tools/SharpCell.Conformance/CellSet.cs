using System;
using System.Collections.Generic;
using System.Linq;
using SharpCell.Evaluation;

namespace SharpCell.Conformance;

/// <summary>
/// Cell addresses of one workbook written compactly per sheet: runs of rows in a column become
/// ranges, as in <c>"A1:A40 C3"</c>, so the report can record exactly which cells match.
/// </summary>
internal static class CellSet
{
    public static SortedDictionary<string, string> Encode(IEnumerable<(string Sheet, CellAddress Cell)> cells)
    {
        var result = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var sheet in cells.GroupBy(c => c.Sheet, StringComparer.Ordinal))
        {
            var tokens = new List<string>();
            foreach (var column in sheet.GroupBy(c => c.Cell.Column).OrderBy(g => g.Key))
            {
                var rows = column.Select(c => c.Cell.Row).Distinct().Order().ToList();
                for (var i = 0; i < rows.Count;)
                {
                    var j = i;
                    while (j + 1 < rows.Count && rows[j + 1] == rows[j] + 1)
                        j++;
                    var first = new CellAddress(rows[i], column.Key).ToString();
                    tokens.Add(i == j ? first : first + ":" + new CellAddress(rows[j], column.Key));
                    i = j + 1;
                }
            }

            result[sheet.Key] = string.Join(' ', tokens);
        }

        return result;
    }

    public static HashSet<(string Sheet, CellAddress Cell)> Decode(IEnumerable<KeyValuePair<string, string>> sheets)
    {
        var result = new HashSet<(string, CellAddress)>();
        foreach (var (sheet, text) in sheets)
        {
            foreach (var token in text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            {
                var colon = token.IndexOf(':');
                if (!CellAddress.TryParse(colon < 0 ? token : token[..colon], out var first)
                    || !CellAddress.TryParse(colon < 0 ? token : token[(colon + 1)..], out var last))
                    throw new FormatException($"'{token}' is not a cell or a column range.");
                var area = new Area(first.Row, first.Column, last.Row, last.Column);
                for (var row = area.FirstRow; row <= area.LastRow; row++)
                {
                    for (var column = area.FirstColumn; column <= area.LastColumn; column++)
                        result.Add((sheet, new CellAddress(row, column)));
                }
            }
        }

        return result;
    }
}
