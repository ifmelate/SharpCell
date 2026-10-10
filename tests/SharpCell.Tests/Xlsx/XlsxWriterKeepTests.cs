using System.IO;
using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

public class XlsxWriterKeepTests
{
    private static readonly XlsxWriteOptions Keep = new() { KeepUncalculated = true };

    private static TestXlsx WithUnknownFunction(string tail = "") => new TestXlsx { WorkbookTail = tail }.Sheet("S",
        "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"
        + "<row r=\"2\"><c r=\"A2\"><f>NOSUCHFN(A1)</f><v>7</v></c></row>"
        + "<row r=\"3\"><c r=\"A3\"><f>A1*2</f><v>2</v></c></row>");

    [Fact]
    public void A_formula_SharpCell_cannot_calculate_keeps_Excels_result()
    {
        var workbook = Load(WithUnknownFunction());
        workbook["S"]["A1"].Value = 5;
        workbook.Recalculate();

        var file = Save(workbook, Keep);
        var sheet = Part(file, "xl/worksheets/sheet1.xml");

        Assert.Contains("<c r=\"A2\"><f>NOSUCHFN(A1)</f><v>7</v></c>", sheet);
        Assert.Contains("<c r=\"A3\"><f>A1*2</f><v>10</v></c>", sheet);
        Assert.Contains("</sheets><calcPr fullCalcOnLoad=\"1\" /></workbook>", Part(file, "xl/workbook.xml"));
    }

    [Fact]
    public void An_existing_calcPr_keeps_its_attributes()
    {
        var workbook = Load(WithUnknownFunction("<calcPr calcId=\"181029\"/>"));

        Assert.Contains("<calcPr calcId=\"181029\" fullCalcOnLoad=\"1\" />", Part(Save(workbook, Keep), "xl/workbook.xml"));
    }

    [Fact]
    public void Without_kept_formulas_the_workbook_part_is_untouched()
    {
        var original = XlsxSourceTests.Template().Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook["S"]["A1"].Value = 5;
        workbook.Recalculate();

        Assert.Equal(PartBytes(original, "xl/workbook.xml"), PartBytes(Save(workbook, Keep), "xl/workbook.xml"));
    }

    // In Excel a spilled cell cannot be cleared on its own, but a value typed over one is the user's.
    [Fact]
    public void Values_typed_inside_a_kept_spill_are_saved()
    {
        var xlsx = new TestXlsx { Metadata = TestXlsx.DynamicArrayMetadata }.Sheet("S",
            "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1:A3\">NOSUCHFN()</f><v>1</v></c></row>"
            + "<row r=\"2\"><c r=\"A2\"><v>2</v></c></row><row r=\"3\"><c r=\"A3\"><v>3</v></c></row>");
        var workbook = Load(xlsx);
        workbook["S"]["A3"].Value = 42;
        workbook["S"]["A4"].Value = 7;
        workbook.Recalculate();

        var sheet = Part(Save(workbook, Keep), "xl/worksheets/sheet1.xml");

        Assert.Contains("<c r=\"A2\"><v>2</v></c>", sheet);
        Assert.Contains("<c r=\"A3\"><v>42</v></c>", sheet);
        Assert.Contains("<c r=\"A4\"><v>7</v></c>", sheet);
    }

    [Fact]
    public void A_kept_dynamic_array_keeps_its_spill()
    {
        var xlsx = new TestXlsx { Metadata = TestXlsx.DynamicArrayMetadata }.Sheet("S",
            "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1:A2\">NOSUCHFN()</f><v>1</v></c><c r=\"B1\"><v>0</v></c></row>"
            + "<row r=\"2\"><c r=\"A2\"><v>2</v></c></row>");
        var workbook = Load(xlsx);
        workbook["S"]["B1"].Value = 1;
        workbook.Recalculate();

        var sheet = Part(Save(workbook, Keep), "xl/worksheets/sheet1.xml");

        Assert.Contains("<c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1:A2\">NOSUCHFN()</f><v>1</v></c>", sheet);
        Assert.Contains("<c r=\"A2\"><v>2</v></c>", sheet);
    }
}
