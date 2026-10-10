using System;
using System.IO;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class XlsxSourceTests
{
    internal static TestXlsx Template() => new TestXlsx().Sheet("S",
        "<row r=\"1\"><c r=\"A1\"><v>2</v></c></row><row r=\"2\"><c r=\"A2\"><f>A1*3</f><v>6</v></c></row>");

    [Fact]
    public void A_loaded_workbook_keeps_the_file_and_its_formulas()
    {
        var bytes = Template().Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(bytes));

        var source = Assert.IsType<XlsxSource>(workbook.Source);
        Assert.Equal(bytes, source.Bytes);
        Assert.Equal("xl/workbook.xml", source.WorkbookPart);
        var sheet = source.Sheets[workbook["S"]];
        Assert.Equal("xl/worksheets/sheet1.xml", sheet.Part);
        var formula = Assert.Single(sheet.Formulas);
        Assert.Equal(new CellAddress(2, 1), formula.Key);
        Assert.Equal("=A1*3", formula.Value.Text);
        Assert.True(formula.Value.IsLegacy);   // saved without the dynamic array flag
        Assert.Null(formula.Value.FixedArray);
    }

    [Fact]
    public void A_workbook_from_a_stream_that_cannot_seek_keeps_the_file()
    {
        var bytes = Template().Build().ToArray();
        var workbook = XlsxReader.Load(new ForwardOnlyStream(bytes));

        Assert.Equal(bytes, Assert.IsType<XlsxSource>(workbook.Source).Bytes);
    }

    [Fact]
    public void A_workbook_from_a_path_keeps_the_file()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        var bytes = Template().Build().ToArray();
        File.WriteAllBytes(path, bytes);
        try
        {
            Assert.Equal(bytes, Assert.IsType<XlsxSource>(XlsxReader.Load(path).Source).Bytes);
        }
        finally
        {
            File.Delete(path);
        }
    }

    public static TheoryData<string> Changes() =>
        ["add sheet", "define name", "hide row", "filter", "date system", "add table"];

    internal static void Change(Workbook workbook, string change)
    {
        var sheet = workbook["S"];
        switch (change)
        {
            case "add sheet": workbook.AddSheet("T"); break;
            case "define name": workbook.DefineName("Rate", "=0.1"); break;
            case "hide row": sheet.SetRowHidden(5, true); break;
            case "filter": sheet.FilterMode = true; break;
            case "date system": workbook.DateSystem = DateSystem.Date1904; break;
            case "add table": sheet.AddTable("Sales", "C1:D3"); break;
        }
    }

    [Theory]
    [MemberData(nameof(Changes))]
    public void The_structure_text_changes_with_the_structure(string change)
    {
        var workbook = XlsxReader.Load(Template().Build());
        var before = Structure.Of(workbook);
        Change(workbook, change);

        Assert.NotEqual(before, Structure.Of(workbook));
    }

    [Fact]
    public void The_structure_text_ignores_cell_values()
    {
        var workbook = XlsxReader.Load(Template().Build());
        var before = Structure.Of(workbook);
        workbook["S"]["A1"].Value = 5;
        workbook.Recalculate();

        Assert.Equal(before, Structure.Of(workbook));
    }

    private sealed class ForwardOnlyStream(byte[] bytes) : MemoryStream(bytes)
    {
        public override bool CanSeek => false;
    }
}
