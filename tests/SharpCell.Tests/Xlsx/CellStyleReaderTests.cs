using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class CellStyleReaderTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string Styles = $$"""
        <styleSheet xmlns="{{Main}}">
          <fonts count="4">
            <font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
            <font><b/><i/><u/><strike/><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/></font>
            <font><b val="0"/><u val="double"/><sz val="9"/><color indexed="12"/><name val="Consolas"/></font>
            <font><u val="singleAccounting"/><color theme="4" tint="-0.499984740745262"/><name val="Calibri"/></font>
          </fonts>
          <fills count="5">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="gray125"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>
            <fill><patternFill patternType="solid"><fgColor theme="0" tint="-0.149998474074526"/></patternFill></fill>
            <fill><patternFill patternType="solid"><fgColor indexed="10"/></patternFill></fill>
          </fills>
          <borders count="2">
            <border><left/><right/><top/><bottom/><diagonal/></border>
            <border><left style="thin"><color indexed="64"/></left><right style="medium"><color rgb="FF0000FF"/></right><top style="dashDotDot"/><bottom style="double"><color auto="1"/></bottom><diagonal/></border>
          </borders>
          <cellXfs count="7">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
            <xf numFmtId="4" fontId="1" fillId="2" borderId="1"/>
            <xf numFmtId="0" fontId="2" fillId="1" borderId="0"><alignment horizontal="center" vertical="top" wrapText="1" indent="3"/></xf>
            <xf numFmtId="0" fontId="3" fillId="3" borderId="0"><alignment horizontal="centerContinuous" vertical="center"/></xf>
            <xf numFmtId="0" fontId="9" fillId="9" borderId="9"><alignment horizontal="nonsense" indent="999"/></xf>
            <xf numFmtId="0" fontId="0" fillId="4" borderId="0"/>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          </cellXfs>
        </styleSheet>
        """;

    private static Worksheet Load(string cells, string? styles = Styles, string? theme = null)
    {
        var file = new TestXlsx { Styles = styles, Theme = theme }.Sheet("S", $"<row r=\"1\">{cells}</row>");
        return XlsxReader.Load(file.Build())["S"];
    }

    private static CellStyle StyleOf(int s, string? styles = Styles, string? theme = null) =>
        Load($"<c r=\"A1\" s=\"{s}\"><v>1</v></c>", styles, theme)["A1"].Style;

    [Fact]
    public void Fonts_are_read_with_their_effects()
    {
        var font = StyleOf(1).Font;

        Assert.Equal("Arial", font.Name);
        Assert.Equal(14, font.Size);
        Assert.True(font.Bold);
        Assert.True(font.Italic);
        Assert.Equal(CellUnderline.Single, font.Underline);
        Assert.True(font.Strikethrough);
        Assert.Equal(CellColor.FromRgb(0xFF0000), font.Color);
    }

    [Fact]
    public void A_false_val_switches_an_effect_off_and_underline_kinds_are_read()
    {
        var font = StyleOf(2).Font;

        Assert.False(font.Bold);
        Assert.Equal(CellUnderline.Double, font.Underline);
        Assert.Equal(9, font.Size);
        Assert.Equal(CellColor.FromRgb(0x0000FF), font.Color);   // indexed 12 in the default palette
        Assert.Equal(CellUnderline.SingleAccounting, StyleOf(3).Font.Underline);
    }

    [Fact]
    public void Solid_fills_are_read_and_patterns_are_not()
    {
        Assert.Equal(CellColor.FromRgb(0xFFFF00), StyleOf(1).Fill);
        Assert.Null(StyleOf(2).Fill);          // gray125
        Assert.Null(StyleOf(0).Fill);
        Assert.Equal(CellColor.FromRgb(0xFF0000), StyleOf(5).Fill);   // indexed 10
    }

    [Fact]
    public void Borders_are_read_edge_by_edge()
    {
        var style = StyleOf(1);

        Assert.Equal(new CellBorder(CellBorderStyle.Thin, CellColor.FromRgb(0x000000)), style.LeftBorder);
        Assert.Equal(new CellBorder(CellBorderStyle.Medium, CellColor.FromRgb(0x0000FF)), style.RightBorder);
        Assert.Equal(new CellBorder(CellBorderStyle.DashDotDot), style.TopBorder);
        Assert.Equal(new CellBorder(CellBorderStyle.Double), style.BottomBorder);
        Assert.Null(StyleOf(0).LeftBorder);
    }

    [Fact]
    public void Alignment_wrapping_and_indent_are_read()
    {
        var style = StyleOf(2);

        Assert.Equal(CellHorizontalAlignment.Center, style.HorizontalAlignment);
        Assert.Equal(CellVerticalAlignment.Top, style.VerticalAlignment);
        Assert.True(style.WrapText);
        Assert.Equal(3, style.Indent);
        Assert.Equal(CellHorizontalAlignment.CenterContinuous, StyleOf(3).HorizontalAlignment);
        Assert.Equal(CellVerticalAlignment.Center, StyleOf(3).VerticalAlignment);
    }

    [Fact]
    public void Theme_colours_come_from_the_theme_part_with_Excels_index_order()
    {
        // Excel numbers theme colours lt1, dk1, lt2, dk2, accent1... while the part lists dk1 first.
        var theme = Theme(dk1: "112233", lt1: "FAFAFA", accent1: "4472C4");

        Assert.Equal(CellColor.FromRgb(0x112233), StyleOf(0, theme: theme).Font.Color);
        Assert.NotNull(StyleOf(3, theme: theme).Font.Color);
    }

    [Fact]
    public void A_tint_moves_the_lightness_in_hue_lightness_and_saturation()
    {
        // White darkened by 15% and black lightened by 50%, as in Excel's colour picker.
        var theme = Theme(dk1: "000000", lt1: "FFFFFF", accent1: "4472C4");
        var lighterBlack = Styles.Replace("<color theme=\"1\"/>", "<color theme=\"1\" tint=\"0.499984740745262\"/>");

        Assert.Equal(CellColor.FromRgb(0xD9D9D9), StyleOf(3, theme: theme).Fill);
        Assert.Equal(CellColor.FromRgb(0x7F7F7F), StyleOf(0, lighterBlack, theme).Font.Color);
    }

    [Fact]
    public void Without_a_theme_part_the_Office_theme_is_used()
    {
        Assert.Equal(CellColor.FromRgb(0x000000), StyleOf(0).Font.Color);
    }

    [Fact]
    public void Indexed_colours_of_the_file_replace_the_default_palette()
    {
        var styles = Styles.Replace("</styleSheet>",
            "<colors><indexedColors>" + string.Concat(Enumerable.Range(0, 13).Select(i => i == 12 ? "<rgbColor rgb=\"FF00AA00\"/>" : "<rgbColor rgb=\"FF000000\"/>")) + "</indexedColors></colors></styleSheet>");

        Assert.Equal(CellColor.FromRgb(0x00AA00), StyleOf(2, styles).Font.Color);
    }

    [Fact]
    public void Ids_a_file_does_not_have_fall_back_to_its_first_entries()
    {
        var style = StyleOf(4);

        Assert.Equal(StyleOf(0).Font, style.Font);
        Assert.Null(style.Fill);
        Assert.Null(style.LeftBorder);
        Assert.Equal(CellHorizontalAlignment.General, style.HorizontalAlignment);
        Assert.Equal(0, style.Indent);
    }

    [Fact]
    public void The_first_cell_style_is_the_workbook_default_and_cells_with_it_have_none_of_their_own()
    {
        var sheet = Load("<c r=\"A1\" s=\"0\"><v>1</v></c><c r=\"B1\" s=\"6\"><v>1</v></c><c r=\"C1\"><v>1</v></c><c r=\"D1\" s=\"1\"/>");

        Assert.Equal(sheet.Workbook.DefaultStyle, sheet["A1"].Style);
        Assert.Same(sheet.Workbook.DefaultStyle, sheet["B1"].Style);
        Assert.Same(sheet.Workbook.DefaultStyle, sheet["C1"].Style);
        Assert.Equal("Arial", sheet["D1"].Style.Font.Name);   // an empty cell keeps its style
        Assert.Equal(CellValue.Empty, sheet["D1"].Value);
        Assert.Equal("#,##0.00", sheet["D1"].NumberFormat);
    }

    [Fact]
    public void No_styles_part_gives_the_default_style()
    {
        var sheet = Load("<c r=\"A1\" s=\"3\"><v>1</v></c>", styles: null);

        Assert.Equal(CellStyle.Default, sheet["A1"].Style);
        Assert.Equal(CellStyle.Default, sheet.Workbook.DefaultStyle);
    }

    [Theory]
    [InlineData("garbage")]
    [InlineData($"<styleSheet xmlns=\"{Main}\"><fonts><font><sz val=\"0\"/><name val=\"\"/></font></fonts><cellXfs><xf fontId=\"0\"/></cellXfs></styleSheet>")]
    [InlineData($"<styleSheet xmlns=\"{Main}\"><fonts><font><sz val=\"big\"/><color rgb=\"zz\" theme=\"99\"/></font></fonts><cellXfs><xf fontId=\"0\"/></cellXfs></styleSheet>")]
    public void A_broken_style_never_stops_a_file(string styles)
    {
        var sheet = Load("<c r=\"A1\" s=\"0\"><v>7</v></c>", styles);

        Assert.Equal(CellValue.Number(7), sheet["A1"].Value);
        Assert.NotNull(sheet["A1"].Style.Font.Name);
    }

    [Fact]
    public void Every_corpus_file_loads_with_its_styles()
    {
        var styled = 0;
        foreach (var relative in CorpusFiles.All)
        {
            Workbook workbook;
            try
            {
                workbook = XlsxReader.Load(CorpusFiles.PathOf(relative));
            }
            catch (NotSupportedException)
            {
                continue;
            }

            styled += workbook.Sheets.Sum(s => s.Styles.Count);
        }

        Assert.True(styled > 0, "No corpus cell has a style of its own.");
    }

    private static string Theme(string dk1, string lt1, string accent1) => $"""
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test">
          <a:themeElements><a:clrScheme name="Test">
            <a:dk1><a:sysClr val="windowText" lastClr="{dk1}"/></a:dk1>
            <a:lt1><a:srgbClr val="{lt1}"/></a:lt1>
            <a:dk2><a:srgbClr val="44546A"/></a:dk2>
            <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
            <a:accent1><a:srgbClr val="{accent1}"/></a:accent1>
            <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
            <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
            <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
            <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
            <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
            <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
            <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
          </a:clrScheme></a:themeElements>
        </a:theme>
        """;
}
