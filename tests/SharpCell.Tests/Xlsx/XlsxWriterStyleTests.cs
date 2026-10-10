using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>Saving copies styles and sheet geometry from the file; changes to them are refused, not lost.</summary>
public class XlsxWriterStyleTests
{
    private const string Styles = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
          <cellXfs count="2"><xf fontId="0"/><xf fontId="1"/></cellXfs>
        </styleSheet>
        """;

    private static Workbook Load() => XlsxWriterTests.Load(new TestXlsx { Styles = Styles }.RawSheet("S", """
        <sheetFormatPr defaultRowHeight="15"/>
        <cols><col min="2" max="2" width="20" customWidth="1"/></cols>
        <sheetData><row r="1" ht="30" customHeight="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row></sheetData>
        <mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells>
        """));

    [Fact]
    public void A_workbook_with_styles_and_geometry_saves_when_nothing_changed()
    {
        var workbook = Load();
        workbook["S"]["A1"].Value = 5;

        var saved = XlsxReader.Load(new MemoryStream(XlsxWriterTests.Save(workbook)))["S"];

        Assert.Equal(CellValue.Number(5), saved["A1"].Value);
        Assert.True(saved["B1"].Style.Font.Bold);
        Assert.Equal(20, saved.ColumnWidth(2));
        Assert.Equal(30, saved.RowHeight(1));
        Assert.Equal("C1:D1", Assert.Single(saved.MergedAreas).Address);
    }

    public static TheoryData<string, Action<Worksheet>> Changes => new()
    {
        { "S!A1", s => s["A1"].Style = CellStyle.Default with { Indent = 1 } },
        { "S!B1", s => s["B1"].Style = null },
        { "column", s => s.SetColumnWidth(2, 25) },
        { "column", s => s.SetColumnHidden(3, true) },
        { "row", s => s.SetRowHeight(1, null) },
        { "row", s => s.DefaultRowHeight = 20 },
        { "merge", s => s.Merge("A3:B3") },
        { "merge", s => s.Unmerge("C1:D1") },
        { "default style", s => s.Workbook.DefaultStyle = CellStyle.Default with { WrapText = true } },
        { "frozen", s => s.FrozenRows = 1 },
        { "frozen", s => s.FrozenColumns = 2 },
        { "gridlines", s => s.ShowGridlines = false },
    };

    [Theory]
    [MemberData(nameof(Changes))]
    public void A_changed_style_or_geometry_is_refused(string mentioned, Action<Worksheet> change)
    {
        var workbook = Load();

        change(workbook["S"]);

        var ex = Assert.Throws<NotSupportedException>(() => XlsxWriterTests.Save(workbook));
        Assert.Contains(mentioned, ex.Message);
    }

    [Fact]
    public void A_style_set_back_to_what_the_file_had_saves()
    {
        var workbook = Load();
        var bold = workbook["S"]["B1"].Style;
        workbook["S"]["B1"].Style = null;
        workbook["S"]["B1"].Style = bold with { };

        Assert.NotEmpty(XlsxWriterTests.Save(workbook));
    }
}
