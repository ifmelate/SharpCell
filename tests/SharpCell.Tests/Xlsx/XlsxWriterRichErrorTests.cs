using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

public class XlsxWriterRichErrorTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    // A1 spills SEQUENCE(3) over A1:A3.
    private static TestXlsx Sequence() => new TestXlsx { Metadata = TestXlsx.DynamicArrayMetadata }.Sheet("S",
        "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1:A3\">_xlfn.SEQUENCE(3)</f><v>1</v></c>"
        + "<c r=\"B1\"><f>A1</f><v>1</v></c></row>"
        + "<row r=\"2\"><c r=\"A2\"><v>2</v></c></row><row r=\"3\"><c r=\"A3\"><v>3</v></c></row>");

    [Fact]
    public void A_blocked_spill_is_saved_as_a_rich_error()
    {
        var workbook = Load(Sequence());
        workbook["S"]["A2"].Value = 9;
        workbook.Recalculate();
        Assert.Equal(ErrorKind.Spill, workbook["S"]["A1"].Value.AsError());

        var file = Save(workbook);
        var sheet = Part(file, Sheet1);

        Assert.Contains("<c r=\"A1\" cm=\"1\" t=\"e\" vm=\"1\"><f t=\"array\" ref=\"A1\">_xlfn.SEQUENCE(3)</f><v>#VALUE!</v></c>", sheet);
        Assert.Contains("<c r=\"B1\" t=\"e\" vm=\"2\"><f>A1</f><v>#VALUE!</v></c>", sheet);
        Assert.Contains("<rv s=\"0\"><v>0</v><v>8</v><v>2</v><v>1</v></rv>", Part(file, "xl/richData/rdrichvalue.xml"));
        // B1 passes A1's error on: Excel marks such a value as propagated (excel-web/spill-blocked.xlsx).
        Assert.Contains("<rv s=\"1\"><v>8</v><v>1</v></rv>", Part(file, "xl/richData/rdrichvalue.xml"));
        Assert.Contains("<k n=\"colOffset\" t=\"i\" /><k n=\"errorType\" t=\"i\" /><k n=\"rwOffset\" t=\"i\" /><k n=\"subType\" t=\"i\" />",
            Part(file, "xl/richData/rdrichvaluestructure.xml"));
        Assert.Contains("<s t=\"_error\"><k n=\"errorType\" t=\"i\" /><k n=\"propagated\" t=\"b\" /></s>",
            Part(file, "xl/richData/rdrichvaluestructure.xml"));
        Assert.Contains("/xl/richData/rdrichvalue.xml", Part(file, "[Content_Types].xml"));
        Assert.Contains("relationships/rdRichValue\"", Part(file, "xl/_rels/workbook.xml.rels"));
        Assert.Contains("rdRichValueTypes", Part(file, "xl/_rels/workbook.xml.rels"));

        var saved = Reload(file)["S"];
        Assert.Equal(ErrorKind.Spill, saved["A1"].Value.AsError());
        Assert.Equal(ErrorKind.Spill, saved["B1"].Value.AsError());
    }

    [Fact]
    public void The_same_error_twice_uses_one_rich_value()
    {
        var workbook = Load(new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>_xlfn._xlws.FILTER(A1:A2,A1:A2&gt;C1)</f><v>1</v></c>"
            + "<c r=\"C1\"><v>0</v></c><c r=\"D1\"><f>_xlfn._xlws.FILTER(A1:A2,A1:A2&gt;C1)</f><v>1</v></c></row>"));
        workbook["S"]["C1"].Value = 100;
        workbook.Recalculate();

        var file = Save(workbook);
        var sheet = Part(file, Sheet1);

        Assert.Contains("<c r=\"B1\" t=\"e\" vm=\"1\">", sheet);
        Assert.Contains("<c r=\"D1\" t=\"e\" vm=\"1\">", sheet);
        Assert.Equal(1, Part(file, "xl/richData/rdrichvalue.xml").Split("<rv ").Length - 1);
        Assert.Equal(ErrorKind.Calc, Reload(file)["S"]["B1"].Value.AsError());
    }

    // A file whose rich value 1 is #CALC! (errorType 13, subType 0), as ERROR.TYPE.xlsx in the corpus.
    private static TestXlsx WithCalcRichValue(string sheetData)
    {
        const string RichData = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
        var metadata = "<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xlrd=\"" + RichData + "\">"
            + "<metadataTypes count=\"1\"><metadataType name=\"XLRICHVALUE\" minSupportedVersion=\"120000\"/></metadataTypes>"
            + "<futureMetadata name=\"XLRICHVALUE\" count=\"1\"><bk><extLst><ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\"><xlrd:rvb i=\"0\"/></ext></extLst></bk></futureMetadata>"
            + "<valueMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></valueMetadata></metadata>";
        return new TestXlsx { Metadata = metadata }
            .Part("xl/richData/rdrichvalue.xml", "<rvData xmlns=\"" + RichData + "\" count=\"1\"><rv s=\"0\"><v>13</v><v>0</v></rv></rvData>")
            .Part("xl/richData/rdrichvaluestructure.xml", "<rvStructures xmlns=\"" + RichData + "\" count=\"1\"><s t=\"_error\"><k n=\"errorType\" t=\"i\"/><k n=\"subType\" t=\"i\"/></s></rvStructures>")
            .Sheet("S", sheetData);
    }

    [Fact]
    public void A_calc_error_read_from_another_cell_is_propagated()
    {
        var workbook = Load(new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>_xlfn._xlws.FILTER(A1,A1&gt;C1)</f><v>1</v></c>"
            + "<c r=\"C1\"><v>0</v></c><c r=\"D1\"><f>B1</f><v>1</v></c></row>"));
        workbook["S"]["C1"].Value = 100;
        workbook.Recalculate();

        var file = Save(workbook);

        Assert.Contains("<c r=\"B1\" t=\"e\" vm=\"1\">", Part(file, Sheet1));
        Assert.Contains("<c r=\"D1\" t=\"e\" vm=\"2\">", Part(file, Sheet1));
        Assert.Contains("<rv s=\"0\"><v>13</v><v>0</v></rv><rv s=\"1\"><v>13</v><v>1</v></rv>", Part(file, "xl/richData/rdrichvalue.xml"));
        Assert.Equal(ErrorKind.Calc, Reload(file)["S"]["D1"].Value.AsError());
    }

    [Fact]
    public void A_rich_value_the_file_has_is_reused()
    {
        var workbook = Load(WithCalcRichValue(
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>_xlfn._xlws.FILTER(A1,A1&gt;C1)</f><v>1</v></c><c r=\"C1\"><v>0</v></c></row>"));
        workbook["S"]["C1"].Value = 100;
        workbook.Recalculate();

        var file = Save(workbook);

        Assert.Contains("<c r=\"B1\" t=\"e\" vm=\"1\">", Part(file, Sheet1));
        Assert.Contains("count=\"1\"", Part(file, "xl/richData/rdrichvalue.xml"));
    }

    // Such values exist in files (ERROR.TYPE.xlsx); only typing one is refused.
    [Fact]
    public void A_saved_calc_value_stays_when_every_value_is_written_again()
    {
        var workbook = Load(WithCalcRichValue(
            "<row r=\"1\"><c r=\"A1\" t=\"e\" vm=\"1\"><v>#VALUE!</v></c><c r=\"B1\"><v>1</v></c></row>"));
        workbook["S"]["B1"].Value = 2;
        workbook.Recalculate();

        var file = Save(workbook, new XlsxWriteOptions { RewriteAllValues = true });

        Assert.Contains("<c r=\"A1\" t=\"e\" vm=\"1\"><v>#VALUE!</v></c>", Part(file, Sheet1));
    }

    [Fact]
    public void Metadata_with_dynamic_arrays_only_gets_the_rich_value_type()
    {
        var workbook = Load(Sequence());
        workbook["S"]["A3"].Value = 9;
        workbook.Recalculate();

        var metadata = Part(Save(workbook), "xl/metadata.xml");

        Assert.Contains("<metadataType name=\"XLRICHVALUE\"", metadata);
        Assert.Contains("<valueMetadata count=\"2\"><bk><rc t=\"2\" v=\"0\" /></bk><bk><rc t=\"2\" v=\"1\" /></bk></valueMetadata>", metadata);
        Assert.Contains("name=\"XLDAPR\"", metadata);
    }
}
