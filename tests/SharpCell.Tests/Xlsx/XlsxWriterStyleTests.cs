using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>Saving writes changed styles and number formats; sheet geometry is copied from the file.</summary>
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
        { "column", s => s.SetColumnWidth(2, 25) },
        { "column", s => s.SetColumnHidden(3, true) },
        { "row", s => s.SetRowHeight(1, null) },
        { "row", s => s.DefaultRowHeight = 20 },
        { "merge", s => s.Merge("A3:B3") },
        { "merge", s => s.Unmerge("C1:D1") },
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

    private const string FullStyles = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
          <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
          <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
            <fill><patternFill patternType="darkGray"><fgColor rgb="FFFF0000"/></patternFill></fill></fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
            <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment textRotation="45"/></xf></cellXfs>
          <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>
        """;

    private static (byte[] Saved, Workbook Reloaded) Edit(Action<Worksheet> change)
    {
        var original = new TestXlsx { Styles = FullStyles }.Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>0.5</v></c><c r=\"B1\" s=\"1\"><v>2</v></c><c r=\"C1\" s=\"2\"><v>3</v></c></row>").Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        change(workbook["S"]);
        var saved = XlsxWriterTests.Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        return (saved, XlsxWriterTests.Reload(saved));
    }

    private static readonly CellStyle Header = CellStyle.Default with
    {
        Font = CellFont.Default with { Bold = true, Italic = true, Underline = CellUnderline.Double, Color = CellColor.FromRgb(0x1F4E79), Size = 14, Name = "Arial" },
        Fill = CellColor.FromRgb(0xD9E1F2),
        BottomBorder = new CellBorder(CellBorderStyle.Thin),
        LeftBorder = new CellBorder(CellBorderStyle.Dashed, CellColor.FromRgb(0xFF0000)),
        HorizontalAlignment = CellHorizontalAlignment.Center,
        VerticalAlignment = CellVerticalAlignment.Top,
        WrapText = true,
        Indent = 2,
    };

    [Fact]
    public void A_style_on_a_new_cell_reads_back_the_same()
    {
        var (_, reloaded) = Edit(s =>
        {
            s["A3"].Value = 1;
            s["A3"].Style = Header;
        });

        Assert.Equal(Header, reloaded["S"]["A3"].Style);
    }

    [Fact]
    public void A_number_format_on_a_loaded_cell_reads_back_the_same()
    {
        var (_, reloaded) = Edit(s => s["A1"].NumberFormat = "0.00%");

        Assert.Equal("0.00%", reloaded["S"]["A1"].NumberFormat);
        Assert.Equal("50.00%", reloaded["S"]["A1"].Text);
    }

    [Fact]
    public void Built_in_formats_use_their_id_and_custom_ones_are_added_once()
    {
        var (saved, reloaded) = Edit(s =>
        {
            s["A2"].Value = 1;
            s["A2"].NumberFormat = "0.00";
            s["B2"].Value = 2;
            s["B2"].NumberFormat = "#,##0.0000";
            s["C2"].Value = 3;
            s["C2"].NumberFormat = "#,##0.0000";
            s["D2"].Value = 4;
            s["D2"].NumberFormat = "0.000";
        });

        var styles = XlsxWriterTests.Part(saved, "xl/styles.xml");
        Assert.Contains("numFmtId=\"2\"", styles);
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(styles, "formatCode=\"#,##0.0000\""));
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(styles, "formatCode=\"0.000\""));
        Assert.Equal(["0.00", "#,##0.0000", "#,##0.0000", "0.000"], new[] { "A2", "B2", "C2", "D2" }.Select(a => reloaded["S"][a].NumberFormat));
    }

    [Fact]
    public void Cells_with_equal_styles_share_one_style_entry()
    {
        var (saved, _) = Edit(s =>
        {
            s["A2"].Value = 1;
            s["A2"].Style = Header;
            s["B2"].Value = 2;
            s["B2"].Style = Header with { };
        });

        var sheet = XlsxWriterTests.Part(saved, "xl/worksheets/sheet1.xml");
        var a2 = System.Text.RegularExpressions.Regex.Match(sheet, "<c r=\"A2\" s=\"(\\d+)\"").Groups[1].Value;
        Assert.NotEmpty(a2);
        Assert.Contains($"<c r=\"B2\" s=\"{a2}\"", sheet);
    }

    [Fact]
    public void A_styled_empty_cell_is_written()
    {
        var (saved, reloaded) = Edit(s => s["B3"].Style = Header);

        var sheet = XlsxWriterTests.Part(saved, "xl/worksheets/sheet1.xml");
        Assert.True(System.Text.RegularExpressions.Regex.IsMatch(sheet, "<c r=\"B3\" s=\"\\d+\" ?/>"), sheet);
        Assert.Equal(Header, reloaded["S"]["B3"].Style);
    }

    [Fact]
    public void A_removed_style_gives_the_default_back()
    {
        var (_, reloaded) = Edit(s => s["B1"].Style = null);

        Assert.False(reloaded["S"]["B1"].Style.Font.Bold);
        Assert.Equal(2, reloaded["S"]["B1"].Value.AsNumber());
    }

    [Fact]
    public void A_changed_default_style_applies_to_cells_without_their_own()
    {
        var wrapped = CellStyle.Default with { WrapText = true };
        var (_, reloaded) = Edit(s => s.Workbook.DefaultStyle = wrapped);

        Assert.Equal(wrapped, reloaded.DefaultStyle);
        Assert.True(reloaded["S"]["A1"].Style.WrapText);
        Assert.True(reloaded["S"]["B1"].Style.Font.Bold);   // its own style stays
    }

    [Fact]
    public void A_changed_format_keeps_what_the_model_does_not_know()
    {
        var (saved, reloaded) = Edit(s => s["C1"].NumberFormat = "0.0");

        var styles = XlsxWriterTests.Part(saved, "xl/styles.xml");
        var index = int.Parse(System.Text.RegularExpressions.Regex.Match(XlsxWriterTests.Part(saved, "xl/worksheets/sheet1.xml"), "<c r=\"C1\" s=\"(\\d+)\"").Groups[1].Value);
        var xf = System.Xml.Linq.XDocument.Parse(styles).Root!.Elements().Single(e => e.Name.LocalName == "cellXfs").Elements().ElementAt(index);
        Assert.Equal("2", (string?)xf.Attribute("fillId"));
        Assert.Equal("45", (string?)xf.Elements().Single().Attribute("textRotation"));
        Assert.Equal("0.0", reloaded["S"]["C1"].NumberFormat);
    }

    [Fact]
    public void A_file_without_styles_gets_a_styles_part()
    {
        var original = new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>").Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        workbook["S"]["A1"].Style = Header;

        var saved = XlsxWriterTests.Save(workbook);

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        Assert.Equal(Header, XlsxWriterTests.Reload(saved)["S"]["A1"].Style);
    }
}
