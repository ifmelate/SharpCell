using SharpCell.Tests.Xlsx;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Model;

/// <summary>Theme colours: made from the workbook's theme, kept as theme colours when saved.</summary>
public class ThemeColorTests
{
    [Fact]
    public void A_new_workbook_has_the_office_theme()
    {
        var workbook = new Workbook();

        Assert.Equal(12, workbook.ThemeColors.Count);
        Assert.Equal(CellColor.FromRgb(0x4472C4), workbook.ThemeColors[(int)ThemeColor.Accent1]);
        Assert.Equal(CellColor.FromRgb(0xFFFFFF), workbook.ThemeColors[(int)ThemeColor.Background1]);
    }

    [Fact]
    public void A_theme_colour_has_the_themes_rgb_with_the_tint()
    {
        var workbook = new Workbook();

        var accent = workbook.ThemeColor(ThemeColor.Accent1);
        var lighter = workbook.ThemeColor(ThemeColor.Background1, -0.5);

        Assert.Equal(CellColor.FromRgb(0x4472C4) with { Theme = ThemeColor.Accent1 }, accent);
        Assert.Equal(CellColor.FromRgb(0x808080) with { Theme = ThemeColor.Background1, Tint = -0.5 }, lighter);
        Assert.NotEqual(CellColor.FromRgb(0x4472C4), accent);
        Assert.Throws<ArgumentOutOfRangeException>(() => workbook.ThemeColor(ThemeColor.Accent1, 1.5));
        Assert.Throws<ArgumentOutOfRangeException>(() => workbook.ThemeColor((ThemeColor)12));
        Assert.Throws<ArgumentOutOfRangeException>(() => CellColor.FromRgb(0) with { Tint = -2 });
    }

    [Fact]
    public void Theme_colours_are_saved_as_theme_colours_and_read_back()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        var fill = workbook.ThemeColor(ThemeColor.Accent2, 0.79998168889431442);
        sheet["A1"].Value = "x";
        sheet["A1"].Style = CellStyle.Default with { Fill = fill, Font = CellFont.Default with { Color = workbook.ThemeColor(ThemeColor.Text2) } };

        var saved = XlsxWriterTests.Save(workbook);
        var styles = XlsxWriterTests.Part(saved, "xl/styles.xml");
        var reloaded = XlsxWriterTests.Reload(saved)["S"]["A1"].Style;

        Assert.Contains("<fgColor theme=\"5\" tint=\"0.7999816888943144\"", styles);
        Assert.Contains("<color theme=\"3\"", styles);
        Assert.Equal(fill, reloaded.Fill);
        Assert.Equal(ThemeColor.Text2, reloaded.Font.Color!.Value.Theme);
    }

    [Fact]
    public void A_clone_keeps_the_theme_of_the_file()
    {
        var file = new TestXlsx { Theme = CellStyleReaderTests.Theme(dk1: "112233", lt1: "FAFAFA", accent1: "ABCDEF") }.Sheet("S", "");
        var workbook = XlsxReader.Load(file.Build());

        Assert.Equal(CellColor.FromRgb(0xABCDEF), workbook.Clone().ThemeColors[(int)ThemeColor.Accent1]);
        Assert.Equal(CellColor.FromRgb(0xABCDEF) with { Theme = ThemeColor.Accent1 }, workbook.ThemeColor(ThemeColor.Accent1));
    }
}
