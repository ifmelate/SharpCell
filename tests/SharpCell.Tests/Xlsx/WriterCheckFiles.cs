using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>
/// Writes the files in tools/excel-oracle/writer-check for opening in Excel: real workbooks Excel
/// saved, changed and written back by SharpCell. README.md lists the values SharpCell saved in the
/// cells to compare before Excel recalculates. Run on demand:
/// <c>dotnet test --project tests/SharpCell.Tests -f net10.0 -- --filter-class "SharpCell.Tests.Xlsx.WriterCheckFiles" --explicit only</c>.
/// </summary>
public class WriterCheckFiles
{
    private sealed record Case(string Name, string Source, string What, Action<Workbook> Change, string[] Cells, bool Keep = false);

    private static readonly Case[] Cases =
    [
        new("inputs-and-text", "../../samples/SharpCell.Sample/budget.xlsx",
            "Rent B2 changed to 1350; text constants in G2:G5 (escape-like text, CR LF, a control character, leading spaces); a new row 20.",
            workbook =>
            {
                var budget = workbook["Budget"];
                budget["B2"].Value = 1350;
                budget["G2"].Value = "_x0041_";
                budget["G3"].Value = "line\r\nbreak";
                budget["G4"].Value = "\u0001 control";
                budget["G5"].Value = "  leading spaces";
                budget["A20"].Value = 42;
            },
            ["Budget!B2", "Budget!B6", "Budget!F2", "Budget!G2", "Budget!G3", "Budget!G4", "Budget!G5", "Budget!A20"]),
        new("spill-grow", "ironcalc/templates/invoice.xlsx",
            "A new item in the first empty cell of C15:C25: the item numbers in column B (SEQUENCE over COUNTA) grow by one.",
            workbook => FirstEmpty(workbook, out _, out _).Value = "Extra item",
            ["B15", "B16", "B17", "B18", "B19", "C18"]),
        new("spill-shrink", "ironcalc/templates/invoice.xlsx",
            "The last item of C15:C25 cleared: the item numbers in column B shrink by one.",
            workbook =>
            {
                FirstEmpty(workbook, out var sheet, out var row);
                sheet[row - 1, 3].Value = CellValue.Empty;
            },
            ["B15", "B16", "B17", "B18", "C17"]),
        new("spill-blocked", "ironcalc/DynamicArrays.xlsx",
            "B5 typed over the SEQUENCE spill at A3:B12: A3 is #SPILL! (a blocked spill, new rich value, the file had no xl/richData); F3 and K3 read it.",
            workbook => workbook["DynamicArrays"]["B5"].Value = 999,
            ["DynamicArrays!A3", "DynamicArrays!F3", "DynamicArrays!K3", "DynamicArrays!B5"]),
        new("rich-errors-cleared", "ironcalc/calc_tests/INFORMATION/ISREF.xlsx",
            "M1 set to TRUE: D7 (was #CALC!) and D9 (was #SPILL!) get plain values and lose their vm; D11 shrinks to one cell.",
            workbook => workbook.Sheets[0]["M1"].Value = true,
            ["D7", "D9", "D11", "D12", "D14"]),
        new("keep-uncalculated", "ironcalc/calc_tests/LOGICAL/IFERROR.xlsx",
            "Saved with KeepUncalculated: formulas calling functions SharpCell does not know keep Excel's results; the file asks Excel to recalculate on open.",
            _ => { },
            [], Keep: true),
    ];

    [Fact(Explicit = true)]
    public void Write_the_files_to_check_in_Excel()
    {
        var folder = Path.GetFullPath(Path.Combine(CorpusFiles.Root, "..", "..", "tools", "excel-oracle", "writer-check"));
        Directory.CreateDirectory(folder);
        var readme = new StringBuilder("# Files written by SharpCell, to check in Excel\n\n")
            .Append("Written by `tests/SharpCell.Tests/Xlsx/WriterCheckFiles.cs`. For each file: open it in Excel (or Excel for the web) and note ")
            .Append("whether it opens without repair, whether the cells below show these values before recalculating, and whether ")
            .Append("Formulas → Calculate Workbook changes anything.\n");
        foreach (var check in Cases)
        {
            var workbook = XlsxReader.Load(CorpusFiles.PathOf(check.Source));
            workbook.Recalculate();
            check.Change(workbook);
            workbook.Recalculate();
            var path = Path.Combine(folder, check.Name + ".xlsx");
            XlsxWriter.Save(workbook, path, new XlsxWriteOptions { KeepUncalculated = check.Keep });

            var saved = XlsxReader.Load(path);
            readme.Append($"\n## {check.Name}.xlsx\n\nFrom `{check.Source}`. {check.What}\n\n");
            foreach (var cell in check.Cells)
                readme.Append($"- `{cell}`: {Show(Find(saved, cell).Value)}\n");
        }

        File.WriteAllText(Path.Combine(folder, "README.md"), readme.ToString());
    }

    // The invoice lists items in C15:C25 of its first sheet.
    private static Cell FirstEmpty(Workbook workbook, out Worksheet sheet, out int row)
    {
        sheet = workbook.Sheets[0];
        for (row = 15; row <= 25; row++)
        {
            if (sheet[row, 3].Value.Kind == CellValueKind.Empty)
                return sheet[row, 3];
        }

        throw new InvalidOperationException("The invoice has no empty item row.");
    }

    private static Cell Find(Workbook workbook, string reference)
    {
        var bang = reference.IndexOf('!');
        var sheet = bang < 0 ? workbook.Sheets[0] : workbook[reference[..bang]];
        return sheet[reference[(bang + 1)..]];
    }

    private static string Show(CellValue value) =>
        value.Kind == CellValueKind.Text ? "\"" + value.AsText().Replace("\r", "\\r").Replace("\n", "\\n").Replace("\u0001", "\\u0001") + "\"" : value.ToString();
}
