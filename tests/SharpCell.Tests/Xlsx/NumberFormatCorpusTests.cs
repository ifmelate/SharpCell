using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>
/// Excel saves no display text, so <c>excel-web/number-formats.xlsx</c> pairs each value shown
/// through a cell style (column C) with <c>TEXT(value, code)</c> (column B) for the same code (D).
/// </summary>
public class NumberFormatCorpusTests
{
    private const string File = "excel-web/number-formats.xlsx";

    [Fact]
    public void Cell_text_matches_Excel_s_TEXT_for_the_cell_format()
    {
        var path = CorpusFiles.PathOf(File);
        Assert.SkipWhen(!System.IO.File.Exists(path), $"{File} is not in the corpus yet.");
        var sheet = XlsxReader.Load(path)["Formats"];   // not recalculated: column B holds Excel's results

        var failures = new List<string>();
        for (var row = 1; sheet[row, 4].Value.Kind == CellValueKind.Text; row++)
        {
            var code = sheet[row, 4].Value.AsText();
            var excel = sheet[row, 2].Value;
            var shown = sheet[row, 3].Text;
            Assert.Equal(code, sheet[row, 3].NumberFormat);
            var expected = excel.Kind == CellValueKind.Text ? excel.AsText() : excel.ToString();

            // TEXT refuses a date code on a number that is no date; a cell shows it as hashes.
            if (excel == CellValue.Error(ErrorKind.Value) && sheet[row, 3].Value.Kind == CellValueKind.Number)
                expected = "#######";
            if (shown != expected)
                failures.Add($"row {row}, {code}: Excel TEXT {expected}, Cell.Text {shown}");
        }

        Assert.True(failures.Count == 0, string.Join("\n", failures));
    }
}
