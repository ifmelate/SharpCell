using System.Collections.Generic;
using System.IO;
using System.Text;
using SharpCell.Parsing;

namespace SharpCell.Tests.Xlsx;

/// <summary>The Excel oracle's case files are well-formed before anyone runs them on Windows.</summary>
public class OracleCaseTests
{
    private static string CasesFolder => Path.Combine(CorpusFiles.Root, "..", "..", "tools", "excel-oracle", "cases");

    public static TheoryData<string> CaseFiles()
    {
        var data = new TheoryData<string>();
        foreach (var path in Directory.EnumerateFiles(CasesFolder, "*.csv"))
            data.Add(Path.GetFileName(path));
        return data;
    }

    [Theory]
    [MemberData(nameof(CaseFiles))]
    public void Every_case_has_four_columns_and_a_formula_SharpCell_parses(string file)
    {
        var rows = ReadCsv(File.ReadAllText(Path.Combine(CasesFolder, file), Encoding.UTF8));
        Assert.Equal(["sheet", "cell", "content", "kind"], rows[0]);
        for (var i = 1; i < rows.Count; i++)
        {
            var row = rows[i];
            Assert.True(row.Count == 4, $"{file} line {i + 1}: {row.Count} columns");
            var (cell, content, kind) = (row[1], row[2], row[3]);
            if (kind.StartsWith("name:"))
            {
                FormulaParser.Parse(content, new CellAddress(1, 1));
                continue;
            }

            var target = kind.StartsWith("array:") ? kind[6..].Split(':')[0] : cell;
            Assert.True(CellAddress.TryParse(target, out var origin), $"{file} line {i + 1}: bad cell '{target}'");
            Assert.True(kind is "" or "legacy" || kind.StartsWith("array:"), $"{file} line {i + 1}: unknown kind '{kind}'");
            if (content.StartsWith('='))
                FormulaParser.Parse(content, origin);
        }
    }

    // RFC 4180: fields separated by commas, quoted fields may contain commas and doubled quotes.
    private static List<List<string>> ReadCsv(string text)
    {
        var rows = new List<List<string>>();
        var row = new List<string>();
        var field = new StringBuilder();
        var quoted = false;
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (quoted)
            {
                if (c == '"' && i + 1 < text.Length && text[i + 1] == '"')
                {
                    field.Append('"');
                    i++;
                }
                else if (c == '"')
                {
                    quoted = false;
                }
                else
                {
                    field.Append(c);
                }
            }
            else if (c == '"')
            {
                quoted = true;
            }
            else if (c == ',')
            {
                row.Add(field.ToString());
                field.Clear();
            }
            else if (c == '\n')
            {
                row.Add(field.ToString().TrimEnd('\r'));
                field.Clear();
                rows.Add(row);
                row = [];
            }
            else
            {
                field.Append(c);
            }
        }

        if (field.Length > 0 || row.Count > 0)
        {
            row.Add(field.ToString());
            rows.Add(row);
        }

        return rows;
    }
}
