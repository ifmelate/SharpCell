namespace SharpCell.Tests.Model;

public class CellStyleTests
{
    private static Worksheet Sheet() => new Workbook().AddSheet("S");

    [Fact]
    public void A_cell_has_the_workbook_default_style_until_it_gets_its_own()
    {
        var sheet = Sheet();

        Assert.Same(sheet.Workbook.DefaultStyle, sheet["A1"].Style);
        Assert.Equal(CellStyle.Default, sheet.Workbook.DefaultStyle);
        Assert.Equal("Calibri", CellStyle.Default.Font.Name);
        Assert.Equal(11, CellStyle.Default.Font.Size);
        Assert.Null(CellStyle.Default.Fill);
        Assert.Equal(CellHorizontalAlignment.General, CellStyle.Default.HorizontalAlignment);
        Assert.Equal(CellVerticalAlignment.Bottom, CellStyle.Default.VerticalAlignment);
    }

    [Fact]
    public void A_style_is_kept_and_null_returns_to_the_default()
    {
        var sheet = Sheet();
        var bold = CellStyle.Default with { Font = CellFont.Default with { Bold = true }, Fill = CellColor.FromRgb(0xFFFF00) };

        sheet["B2"].Style = bold;
        Assert.Equal(bold, sheet["B2"].Style);
        Assert.True(sheet["B2"].Style.Font.Bold);

        sheet["B2"].Style = null;
        Assert.Same(sheet.Workbook.DefaultStyle, sheet["B2"].Style);
    }

    [Fact]
    public void Changing_the_default_style_changes_every_cell_without_its_own()
    {
        var sheet = Sheet();
        var own = CellStyle.Default with { Indent = 2 };
        sheet["A1"].Style = own;
        var arial = CellStyle.Default with { Font = CellFont.Default with { Name = "Arial", Size = 10 } };

        sheet.Workbook.DefaultStyle = arial;

        Assert.Equal(arial, sheet["Z9"].Style);
        Assert.Equal(own, sheet["A1"].Style);
        Assert.Throws<ArgumentNullException>(() => sheet.Workbook.DefaultStyle = null!);
    }

    [Fact]
    public void A_style_alone_is_no_content()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A2"].Style = CellStyle.Default with { WrapText = true };
        sheet["A1"].Formula = "=SEQUENCE(2)";
        sheet["B1"].Formula = "=COUNTA(A2:A9)";
        workbook.Recalculate();

        Assert.Equal(CellValue.Number(2), sheet["A2"].Value);
        Assert.Equal(CellValue.Number(1), sheet["B1"].Value);
        Assert.True(sheet["A2"].Style.WrapText);
        Assert.DoesNotContain(sheet.Cells, c => c.Address == "C3");
    }

    [Fact]
    public void A_clone_has_the_same_styles()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        var italic = CellStyle.Default with { Font = CellFont.Default with { Italic = true } };
        sheet["C3"].Style = italic;
        workbook.DefaultStyle = CellStyle.Default with { Indent = 1 };

        var clone = workbook.Clone();

        Assert.Equal(italic, clone["S"]["C3"].Style);
        Assert.Equal(1, clone.DefaultStyle.Indent);
        clone["S"]["C3"].Style = null;
        Assert.Equal(italic, sheet["C3"].Style);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(251)]
    public void An_indent_outside_Excels_range_is_refused(int indent)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => CellStyle.Default with { Indent = indent });
    }

    [Theory]
    [InlineData(0)]
    [InlineData(410)]
    [InlineData(double.NaN)]
    public void A_font_size_outside_Excels_range_is_refused(double size)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => CellFont.Default with { Size = size });
    }

    [Fact]
    public void A_font_needs_a_name()
    {
        Assert.Throws<ArgumentException>(() => CellFont.Default with { Name = "" });
    }

    [Fact]
    public void Colours_read_as_hex()
    {
        Assert.Equal("#4472C4", CellColor.FromRgb(0x4472C4).ToString());
        Assert.Equal(new CellColor(0x44, 0x72, 0xC4), CellColor.FromRgb(0x4472C4));
    }

    [Fact]
    public void Styles_compare_by_value()
    {
        var a = CellStyle.Default with { LeftBorder = new CellBorder(CellBorderStyle.Thin, CellColor.FromRgb(0)) };
        var b = CellStyle.Default with { LeftBorder = new CellBorder(CellBorderStyle.Thin, CellColor.FromRgb(0)) };

        Assert.Equal(a, b);
        Assert.NotEqual(a, CellStyle.Default);
    }
}
