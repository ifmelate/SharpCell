using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class XlsxWriterTests
{
    internal static Workbook Load(TestXlsx xlsx)
    {
        var workbook = XlsxReader.Load(xlsx.Build());
        workbook.Recalculate();
        return workbook;
    }

    internal static byte[] Save(Workbook workbook, XlsxWriteOptions? options = null)
    {
        var output = new MemoryStream();
        XlsxWriter.Save(workbook, output, options ?? new XlsxWriteOptions());
        return output.ToArray();
    }

    internal static Workbook Reload(byte[] file) => XlsxReader.Load(new MemoryStream(file));   // not recalculated: saved values

    internal static string Part(byte[] file, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(file));
        using var reader = new StreamReader(zip.GetEntry(name)!.Open());
        return reader.ReadToEnd();
    }

    internal static byte[] PartBytes(byte[] file, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(file));
        using var stream = zip.GetEntry(name)!.Open();
        var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    [Fact]
    public void A_changed_input_updates_the_constant_and_the_formula()
    {
        var workbook = Load(XlsxSourceTests.Template());
        workbook["S"]["A1"].Value = 5;
        workbook.Recalculate();

        var file = Save(workbook);

        Assert.Contains("<c r=\"A1\"><v>5</v></c>", Part(file, Sheet1));
        Assert.Contains("<c r=\"A2\"><f>A1*3</f><v>15</v></c>", Part(file, Sheet1));
        Assert.Equal(15, Reload(file)["S"]["A2"].Value.AsNumber());
    }

    [Fact]
    public void Cells_that_did_not_change_are_copied_as_they_are()
    {
        var xlsx = new TestXlsx { SharedStrings = "<si><t>Rent</t></si>" }.Sheet("S",
            "<row r=\"1\" spans=\"1:3\" ht=\"20\" customHeight=\"1\"><c r=\"A1\" s=\"3\" t=\"s\"><v>0</v></c><c r=\"B1\"><v>2</v></c>"
            + "<c r=\"C1\" s=\"1\"><f>B1*2</f><v>4</v></c></row>");
        var workbook = Load(xlsx);
        workbook["S"]["B1"].Value = 3;
        workbook.Recalculate();

        var sheet = Part(Save(workbook), Sheet1);

        Assert.Contains("<row r=\"1\" spans=\"1:3\" ht=\"20\" customHeight=\"1\"><c r=\"A1\" s=\"3\" t=\"s\"><v>0</v></c>", sheet);
        Assert.Contains("<c r=\"C1\" s=\"1\"><f>B1*2</f><v>6</v></c>", sheet);
    }

    [Fact]
    public void Text_booleans_and_errors_are_written_with_their_type()
    {
        var xlsx = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\" t=\"str\"><f>A1&amp;\"x\"</f><v>1x</v></c>"
            + "<c r=\"C1\" t=\"b\"><f>A1&gt;0</f><v>1</v></c><c r=\"D1\"><f>1/A1</f><v>1</v></c></row>");
        var workbook = Load(xlsx);
        workbook["S"]["A1"].Value = 0;
        workbook.Recalculate();

        var file = Save(workbook);
        var sheet = Part(file, Sheet1);

        Assert.Contains("<c r=\"B1\" t=\"str\"><f>A1&amp;\"x\"</f><v>0x</v></c>", sheet);
        Assert.Contains("<c r=\"C1\" t=\"b\"><f>A1&gt;0</f><v>0</v></c>", sheet);
        Assert.Contains("<c r=\"D1\" t=\"e\"><f>1/A1</f><v>#DIV/0!</v></c>", sheet);
        Assert.Equal(ErrorKind.Div0, Reload(file)["S"]["D1"].Value.AsError());
    }

    [Theory]
    [InlineData("_x0041_")]
    [InlineData("line\r\nbreak\ttab")]
    [InlineData("\u0001 control")]
    [InlineData(" spaces ")]
    public void Text_constants_survive_writing_and_reading(string text)
    {
        var workbook = Load(XlsxSourceTests.Template());
        workbook["S"]["A1"].Value = text;
        workbook.Recalculate();

        var file = Save(workbook);

        Assert.Contains("<c r=\"A1\" t=\"inlineStr\"><is><t xml:space=\"preserve\">", Part(file, Sheet1));
        Assert.Equal(text, Reload(file)["S"]["A1"].Value.AsText());
    }

    [Fact]
    public void A_new_cell_goes_between_its_neighbours()
    {
        var xlsx = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"C1\"><v>3</v></c></row><row r=\"3\"><c r=\"A3\"><v>9</v></c></row>");
        var workbook = Load(xlsx);
        workbook["S"]["B1"].Value = 2;
        workbook["S"]["B2"].Value = 5;
        workbook["S"]["A4"].Value = 7;
        workbook.Recalculate();

        var sheet = Part(Save(workbook), Sheet1);

        Assert.Contains("<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\"><v>3</v></c></row>"
            + "<row r=\"2\"><c r=\"B2\"><v>5</v></c></row><row r=\"3\"><c r=\"A3\"><v>9</v></c></row>"
            + "<row r=\"4\"><c r=\"A4\"><v>7</v></c></row>", sheet);
    }

    [Fact]
    public void A_value_in_an_empty_sheet_gets_its_row()
    {
        var workbook = Load(new TestXlsx().RawSheet("S", "<sheetData/>"));
        workbook["S"]["B2"].Value = 1;
        workbook.Recalculate();

        Assert.Contains("<sheetData><row r=\"2\"><c r=\"B2\"><v>1</v></c></row></sheetData>", Part(Save(workbook), Sheet1));
    }

    [Fact]
    public void A_cleared_cell_keeps_only_its_style()
    {
        var xlsx = new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\" s=\"4\"><v>1</v></c><c r=\"B1\"><v>2</v></c></row>");
        var workbook = Load(xlsx);
        workbook["S"]["A1"].Value = CellValue.Empty;
        workbook["S"]["B1"].Value = CellValue.Empty;
        workbook.Recalculate();

        Assert.Contains("<row r=\"1\"><c r=\"A1\" s=\"4\" /></row>", Part(Save(workbook), Sheet1));
    }

    [Fact]
    public void Cells_without_positions_keep_theirs_when_a_cell_is_added()
    {
        var xlsx = new TestXlsx().Sheet("S",
            "<row><c><v>1</v></c><c><v>2</v></c></row><row><c><f>A1+B1</f><v>3</v></c></row>");
        var workbook = Load(xlsx);
        workbook["S"]["A1"].Value = 10;
        workbook["S"]["C1"].Value = 5;
        workbook.Recalculate();

        var saved = Reload(Save(workbook))["S"];

        Assert.Equal(10, saved["A1"].Value.AsNumber());
        Assert.Equal(2, saved["B1"].Value.AsNumber());
        Assert.Equal(5, saved["C1"].Value.AsNumber());
        Assert.Equal(12, saved["A2"].Value.AsNumber());
    }

    [Fact]
    public void Only_changed_sheets_are_written_again()
    {
        var xlsx = XlsxSourceTests.Template().Sheet("T", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>")
            .Part("xl/custom.xml", "<anything kept=\"yes\"/>");
        var original = xlsx.Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        workbook["S"]["A1"].Value = 5;
        workbook.Recalculate();

        var file = Save(workbook);

        Assert.NotEqual(PartBytes(original, Sheet1), PartBytes(file, Sheet1));
        foreach (var name in new[] { "xl/worksheets/sheet2.xml", "xl/workbook.xml", "xl/custom.xml", "[Content_Types].xml" })
            Assert.Equal(PartBytes(original, name), PartBytes(file, name));
        using var before = new ZipArchive(new MemoryStream(original));
        using var after = new ZipArchive(new MemoryStream(file));
        Assert.Equal(before.Entries.Select(e => e.FullName), after.Entries.Select(e => e.FullName));
    }

    [Fact]
    public void Saving_twice_gives_the_same_file()
    {
        var workbook = Load(XlsxSourceTests.Template());
        workbook["S"]["A1"].Value = 5;
        workbook.Recalculate();

        Assert.Equal(Save(workbook), Save(workbook));
    }

    [Fact]
    public void The_file_a_workbook_was_read_from_can_be_saved_over()
    {
        var folder = Directory.CreateTempSubdirectory().FullName;
        try
        {
            var path = Path.Combine(folder, "template.xlsx");
            File.WriteAllBytes(path, XlsxSourceTests.Template().Build().ToArray());
            var workbook = XlsxReader.Load(path);
            workbook["S"]["A1"].Value = 5;
            workbook.Recalculate();

            XlsxWriter.Save(workbook, path);

            Assert.Equal(15, XlsxReader.Load(path)["S"]["A2"].Value.AsNumber());
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void A_data_table_result_cannot_be_changed()
    {
        var xlsx = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><f t=\"dataTable\" ref=\"A1\" dt2D=\"0\" dtr=\"0\" r1=\"B1\"/><v>4</v></c></row>");
        var workbook = Load(xlsx);
        workbook["S"]["A1"].Value = 5;
        workbook.Recalculate();

        Assert.Throws<NotSupportedException>(() => Save(workbook));
    }

    [Fact]
    public void A_typed_spill_or_calc_error_cannot_be_saved()
    {
        var workbook = Load(XlsxSourceTests.Template());
        workbook["S"]["A1"].Value = CellValue.Error(ErrorKind.Calc);
        workbook.Recalculate();

        Assert.Throws<NotSupportedException>(() => Save(workbook));
    }
}
