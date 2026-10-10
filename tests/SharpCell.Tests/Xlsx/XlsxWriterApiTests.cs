using System;
using System.IO;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class XlsxWriterApiTests
{
    [Fact]
    public void A_workbook_without_changes_is_saved_as_the_original_bytes()
    {
        var bytes = XlsxSourceTests.Template().Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(bytes));
        workbook.Recalculate();
        var output = new MemoryStream();

        XlsxWriter.Save(workbook, output);

        Assert.Equal(bytes, output.ToArray());
        Assert.True(output.CanWrite);   // left open
    }

    [Fact]
    public void Saving_to_a_path_leaves_no_temporary_file()
    {
        var folder = Directory.CreateTempSubdirectory().FullName;
        try
        {
            var workbook = XlsxReader.Load(XlsxSourceTests.Template().Build());
            workbook.Recalculate();
            var path = Path.Combine(folder, "out.xlsx");

            XlsxWriter.Save(workbook, path);

            Assert.Equal([path], Directory.GetFiles(folder));
            Assert.Equal(6, XlsxReader.Load(path)["S"]["A2"].Value.AsNumber());
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void A_refused_save_writes_nothing()
    {
        var folder = Directory.CreateTempSubdirectory().FullName;
        try
        {
            var workbook = XlsxReader.Load(XlsxSourceTests.Template().Build());   // not recalculated
            var output = new MemoryStream();
            var path = Path.Combine(folder, "out.xlsx");

            Assert.Throws<InvalidOperationException>(() => XlsxWriter.Save(workbook, output));
            Assert.Throws<InvalidOperationException>(() => XlsxWriter.Save(workbook, path));

            Assert.Equal(0, output.Length);
            Assert.Empty(Directory.GetFiles(folder));
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void Arguments_are_checked()
    {
        var workbook = new Workbook();
        Assert.Throws<ArgumentNullException>(() => XlsxWriter.Save(null!, new MemoryStream()));
        Assert.Throws<ArgumentNullException>(() => XlsxWriter.Save(workbook, (Stream)null!));
        Assert.Throws<ArgumentNullException>(() => XlsxWriter.Save(workbook, (string)null!));
        Assert.Throws<ArgumentNullException>(() => XlsxWriter.Save(workbook, new MemoryStream(), null!));
    }

    [Theory]
    [InlineData("A1", "0.00")]
    [InlineData("B1", "0.00")]
    [InlineData("C1", "General")]
    public void A_changed_number_format_cannot_be_saved(string cell, string code)
    {
        var file = new TestXlsx
        {
            Styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"10\"/></cellXfs></styleSheet>",
        }.Sheet("S", "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"C1\" s=\"1\"><v>1</v></c></row>");
        var workbook = XlsxWriterTests.Load(file);
        Assert.Equal("0.00%", workbook["S"]["C1"].NumberFormat);
        var unchanged = XlsxWriterTests.Save(workbook);

        workbook["S"][cell].NumberFormat = code;
        var ex = Assert.Throws<NotSupportedException>(() => XlsxWriterTests.Save(workbook));
        Assert.Contains($"S!{cell}", ex.Message);
        Assert.Contains("format", ex.Message);
        Assert.NotEmpty(unchanged);
    }

    [Fact]
    public void A_format_set_back_to_what_the_file_had_saves()
    {
        var workbook = XlsxWriterTests.Load(new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"));
        workbook["S"]["A1"].NumberFormat = "0.00";
        workbook["S"]["A1"].NumberFormat = "General";
        Assert.NotEmpty(XlsxWriterTests.Save(workbook));
    }
}
