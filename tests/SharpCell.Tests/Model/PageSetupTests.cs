namespace SharpCell.Tests.Model;

/// <summary>Paper size, scaling, fitting to pages, and headers and footers.</summary>
public class PageSetupTests
{
    [Fact]
    public void A_new_sheet_prints_on_letter_paper_at_full_size_without_headers()
    {
        var sheet = new Workbook().AddSheet("S");

        Assert.Equal(PaperSize.Letter, sheet.PaperSize);
        Assert.Equal(100, sheet.PrintScale);
        Assert.Null(sheet.FitToPages);
        Assert.Null(sheet.PageHeader);
        Assert.Null(sheet.PageFooter);
        Assert.Null(sheet.FirstPageHeader);
        Assert.Null(sheet.EvenPageFooter);
    }

    [Fact]
    public void Settings_are_set_and_read_back()
    {
        var sheet = new Workbook().AddSheet("S");

        sheet.PaperSize = PaperSize.A4;
        sheet.PrintScale = 85;
        sheet.FitToPages = new PageFit(1, 0);
        sheet.PageHeader = new HeaderFooter(Center: "Sales &A");
        sheet.PageFooter = new HeaderFooter(Right: $"Page {HeaderFooter.PageNumber} of {HeaderFooter.PageCount}");
        sheet.FirstPageHeader = new HeaderFooter(Left: "Cover");

        Assert.Equal((PaperSize.A4, 85, new PageFit(1, 0)), (sheet.PaperSize, sheet.PrintScale, sheet.FitToPages));
        Assert.Equal("Sales &A", sheet.PageHeader!.Center);
        Assert.Equal("Page &P of &N", sheet.PageFooter!.Right);
        Assert.Equal("Cover", sheet.FirstPageHeader!.Left);
        Assert.Null(sheet.FirstPageFooter);
    }

    [Fact]
    public void Wrong_settings_are_refused()
    {
        var sheet = new Workbook().AddSheet("S");

        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.PrintScale = 9);
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.PrintScale = 401);
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.PaperSize = 0);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageFit(-1, 1));
        Assert.Throws<ArgumentException>(() => sheet.PageHeader = new HeaderFooter(Center: new string('x', 254)));
        sheet.PageHeader = new HeaderFooter(Center: new string('x', 253));   // with &C, Excel's 255 exactly
    }

    [Theory]
    [InlineData("&LLeft&CMiddle&RRight", "Left", "Middle", "Right")]
    [InlineData("Plain", null, "Plain", null)]
    [InlineData("&R&P", null, null, "&P")]
    [InlineData("&L&\"Arial,Bold\"&12Title &&Co&R&D", "&\"Arial,Bold\"&12Title &&Co", null, "&D")]
    [InlineData("&\"Arial,&LBold\"x", null, "&\"Arial,&LBold\"x", null)]
    [InlineData("&La&Lb", "ab", null, null)]
    [InlineData("&RRight&LLeft", "Left", null, "Right")]
    [InlineData(" spaced &L x ", " x ", " spaced ", null)]
    public void Excel_codes_are_split_into_sections(string code, string? left, string? center, string? right)
    {
        var text = HeaderFooter.Parse(code);

        Assert.Equal((left, center, right), (text.Left, text.Center, text.Right));
    }

    [Fact]
    public void Sections_are_joined_into_excel_codes()
    {
        Assert.Equal("&LLeft&CMiddle&RRight", new HeaderFooter("Left", "Middle", "Right").ToString());
        Assert.Equal("&R&P", new HeaderFooter(Right: "&P").ToString());
        Assert.Equal("", new HeaderFooter().ToString());
        Assert.Equal("R&&D", HeaderFooter.Escape("R&D"));
    }

    [Fact]
    public void A_clone_prints_the_same()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet.PaperSize = PaperSize.A3;
        sheet.PrintScale = 50;
        sheet.FitToPages = new PageFit(2, 3);
        sheet.PageFooter = new HeaderFooter(Center: "&P");
        sheet.EvenPageFooter = new HeaderFooter(Left: "&P");

        var clone = workbook.Clone()["S"];

        Assert.Equal((PaperSize.A3, 50, new PageFit(2, 3)), (clone.PaperSize, clone.PrintScale, clone.FitToPages));
        Assert.Equal((sheet.PageFooter, sheet.EvenPageFooter), (clone.PageFooter, clone.EvenPageFooter));
    }
}
