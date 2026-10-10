using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

public class XlsxWriterSpillTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    // A1 spills SEQUENCE(B1) down column A; B1 is 2 in the file.
    private static TestXlsx Sequence(string a2Attributes = "") => new TestXlsx { Metadata = TestXlsx.DynamicArrayMetadata }.RawSheet("S",
        "<dimension ref=\"A1:B2\"/><sheetData>"
        + "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1:A2\">_xlfn.SEQUENCE(B1)</f><v>1</v></c><c r=\"B1\"><v>2</v></c></row>"
        + $"<row r=\"2\"><c r=\"A2\"{a2Attributes}><v>2</v></c></row>"
        + "</sheetData>");

    [Fact]
    public void A_growing_spill_writes_its_new_cells_and_area()
    {
        var workbook = Load(Sequence());
        workbook["S"]["B1"].Value = 3;
        workbook.Recalculate();

        var file = Save(workbook);
        var sheet = Part(file, Sheet1);

        Assert.Contains("<c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1:A3\">_xlfn.SEQUENCE(B1)</f><v>1</v></c>", sheet);
        Assert.Contains("<row r=\"3\"><c r=\"A3\"><v>3</v></c></row>", sheet);
        Assert.Contains("<dimension ref=\"A1:B3\" />", sheet);
        var saved = Reload(file);
        saved.Recalculate();
        Assert.Equal(3, saved["S"]["A3"].Value.AsNumber());
    }

    [Fact]
    public void A_shrinking_spill_clears_the_cells_it_left()
    {
        var workbook = Load(Sequence());
        workbook["S"]["B1"].Value = 1;
        workbook.Recalculate();

        var sheet = Part(Save(workbook), Sheet1);

        Assert.Contains("<f t=\"array\" ref=\"A1\">", sheet);
        Assert.Contains("<row r=\"2\"></row>", sheet);
        Assert.Contains("<dimension ref=\"A1:B2\" />", sheet);   // not narrowed
    }

    [Fact]
    public void A_styled_cell_a_spill_left_keeps_its_style()
    {
        var workbook = Load(Sequence(" s=\"5\""));
        workbook["S"]["B1"].Value = 1;
        workbook.Recalculate();

        Assert.Contains("<row r=\"2\"><c r=\"A2\" s=\"5\" /></row>", Part(Save(workbook), Sheet1));
    }

    [Fact]
    public void An_array_formula_keeps_its_area()
    {
        var xlsx = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:A2\">B1:B2*2</f><v>2</v></c><c r=\"B1\"><v>1</v></c></row>"
            + "<row r=\"2\"><c r=\"A2\"><v>4</v></c><c r=\"B2\"><v>2</v></c></row>");
        var workbook = Load(xlsx);
        workbook["S"]["B2"].Value = 5;
        workbook.Recalculate();

        var sheet = Part(Save(workbook), Sheet1);

        Assert.Contains("<f t=\"array\" ref=\"A1:A2\">B1:B2*2</f>", sheet);
        Assert.Contains("<c r=\"A2\"><v>10</v></c>", sheet);
    }
}
