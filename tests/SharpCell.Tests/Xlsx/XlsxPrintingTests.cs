using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

/// <summary>Page setup, print areas, print titles and page breaks are read from files and saved back.</summary>
public class XlsxPrintingTests
{
    private static TestXlsx Book() => new TestXlsx()
        .RawSheet("My Sheet", """
            <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
            <pageMargins left="0.25" right="0.25" top="1" bottom="1" header="0.5" footer="0.5"/>
            <pageSetup paperSize="9" orientation="landscape" r:id="rP"/>
            <rowBreaks count="2" manualBreakCount="2"><brk id="10" max="16383" man="1"/><brk id="20" max="16383" man="1"/></rowBreaks>
            <colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"/></colBreaks>
            """)
        .Sheet("Plain", "")
        .Name("_xlnm.Print_Area", "'My Sheet'!$A$1:$F$40,'My Sheet'!$H$1:$J$5", " localSheetId=\"0\"")
        .Name("_xlnm.Print_Titles", "'My Sheet'!$A:$B,'My Sheet'!$1:$2", " localSheetId=\"0\"");

    private static (byte[] Original, Workbook Workbook) Load()
    {
        var original = Book().Build().ToArray();
        return (original, XlsxReader.Load(new MemoryStream(original)));
    }

    private static (string Sheet, string Workbook, Workbook Reloaded) SaveChecked(byte[] original, Workbook workbook)
    {
        var saved = Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var reloaded = Reload(saved);
        foreach (var sheet in workbook.Sheets)
            Assert.Equal(Describe(sheet), Describe(reloaded[sheet.Name]));
        return (Part(saved, "xl/worksheets/sheet1.xml"), Part(saved, "xl/workbook.xml"), reloaded);
    }

    private static string Describe(Worksheet sheet) =>
        $"{sheet.PageOrientation} {sheet.PageMargins} {sheet.PrintArea} {sheet.PrintTitleRows} {sheet.PrintTitleColumns} "
        + $"{string.Join(",", sheet.RowPageBreaks)} {string.Join(",", sheet.ColumnPageBreaks)} "
        + $"{sheet.PaperSize} {sheet.PrintScale} {sheet.FitToPages} {sheet.PageHeader} {sheet.PageFooter} {sheet.EvenPageHeader} "
        + $"{sheet.EvenPageFooter} {sheet.FirstPageHeader} {sheet.FirstPageFooter}";

    [Fact]
    public void Print_settings_are_read()
    {
        var (_, workbook) = Load();
        var sheet = workbook["My Sheet"];

        Assert.Equal(PageOrientation.Landscape, sheet.PageOrientation);
        Assert.Equal(new PageMargins(0.25, 0.25, 1, 1, 0.5, 0.5), sheet.PageMargins);
        Assert.Equal("A1:F40,H1:J5", sheet.PrintArea);
        Assert.Equal(("1:2", "A:B"), (sheet.PrintTitleRows, sheet.PrintTitleColumns));
        Assert.Equal([11, 21], sheet.RowPageBreaks);
        Assert.Equal([4], sheet.ColumnPageBreaks);
        Assert.Equal((PageOrientation.Portrait, PageMargins.Normal, null), (workbook["Plain"].PageOrientation, workbook["Plain"].PageMargins, workbook["Plain"].PrintArea));
        Assert.DoesNotContain(workbook.DefinedNames, name => name.Name.Contains("Print", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Unchanged_settings_stay_as_the_file_has_them()
    {
        var (original, workbook) = Load();
        workbook["My Sheet"]["A1"].Value = 2;

        var (sheet, book, _) = SaveChecked(original, workbook);

        Assert.Contains("<pageSetup paperSize=\"9\" orientation=\"landscape\" r:id=\"rP\"", sheet);
        Assert.Contains(">'My Sheet'!$A$1:$F$40,'My Sheet'!$H$1:$J$5</definedName>", book);
    }

    [Fact]
    public void Changed_settings_are_saved()
    {
        var (original, workbook) = Load();
        var sheet = workbook["My Sheet"];
        sheet.PageOrientation = PageOrientation.Portrait;
        sheet.PageMargins = PageMargins.Wide;
        sheet.PrintArea = "B2:C3";
        sheet.PrintTitleColumns = null;
        sheet.RemoveRowPageBreak(11);
        sheet.RemoveColumnPageBreak(4);

        var (part, book, _) = SaveChecked(original, workbook);

        Assert.Contains("<pageSetup paperSize=\"9\" orientation=\"portrait\" r:id=\"rP\"", part);
        Assert.Contains("<pageMargins left=\"1\" right=\"1\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"", part);
        Assert.Contains("<brk id=\"20\"", part);
        Assert.DoesNotContain("<brk id=\"10\"", part);
        Assert.DoesNotContain("colBreaks", part);
        Assert.Contains(">'My Sheet'!$B$2:$C$3</definedName>", book);
        Assert.Contains(">'My Sheet'!$1:$2</definedName>", book);
    }

    [Fact]
    public void Settings_on_a_sheet_without_any_are_added()
    {
        var (original, workbook) = Load();
        var sheet = workbook["Plain"];
        sheet.PageOrientation = PageOrientation.Landscape;
        sheet.PageMargins = PageMargins.Narrow;
        sheet.PrintTitleRows = "1:1";
        sheet.AddRowPageBreak(30);
        workbook["My Sheet"].PrintArea = null;
        workbook["My Sheet"].PrintTitleRows = null;
        workbook["My Sheet"].PrintTitleColumns = null;

        var (_, book, reloaded) = SaveChecked(original, workbook);

        Assert.Equal((PageOrientation.Landscape, "1:1", 30), (reloaded["Plain"].PageOrientation, reloaded["Plain"].PrintTitleRows, reloaded["Plain"].RowPageBreaks.Single()));
        Assert.DoesNotContain("'My Sheet'!", book);
        Assert.Contains("localSheetId=\"1\">Plain!$1:$1</definedName>", book);
    }

    [Fact]
    public void Settings_follow_inserted_rows_when_saved()
    {
        var (original, workbook) = Load();

        workbook["My Sheet"].InsertRows(1, 2);
        var (part, book, _) = SaveChecked(original, workbook);

        Assert.Contains(">'My Sheet'!$A$3:$F$42,'My Sheet'!$H$3:$J$7</definedName>", book);
        Assert.Contains(">'My Sheet'!$A:$B,'My Sheet'!$3:$4</definedName>", book);
        Assert.Contains("<brk id=\"12\"", part);
    }

    [Fact]
    public void A_workbook_built_in_code_saves_its_print_settings()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Report");
        sheet["A1"].Value = "x";
        sheet.PageOrientation = PageOrientation.Landscape;
        sheet.PageMargins = PageMargins.Narrow;
        sheet.PrintArea = "A1:D20";
        sheet.PrintTitleRows = "1:1";
        sheet.AddRowPageBreak(11);
        sheet.AddColumnPageBreak(3);

        var saved = Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        Assert.Equal(Describe(sheet), Describe(Reload(saved)["Report"]));
    }

    private static (byte[] Original, Workbook Workbook) LoadSetup()
    {
        var original = new TestXlsx()
            .RawSheet("Setup", """
                <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
                <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
                <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                <pageSetup paperSize="9" scale="80" fitToHeight="0" orientation="portrait" r:id="rP"/>
                <headerFooter differentFirst="1" scaleWithDoc="0"><oddHeader>&amp;LSales&amp;R&amp;D</oddHeader><oddFooter>&amp;CPage &amp;P of &amp;N</oddFooter><firstHeader>&amp;CCover</firstHeader></headerFooter>
                """)
            .Sheet("Plain", "")
            .Build().ToArray();
        return (original, XlsxReader.Load(new MemoryStream(original)));
    }

    [Fact]
    public void Paper_scale_fit_and_headers_are_read()
    {
        var (_, workbook) = LoadSetup();
        var sheet = workbook["Setup"];

        Assert.Equal((PaperSize.A4, 80, new PageFit(1, 0)), (sheet.PaperSize, sheet.PrintScale, sheet.FitToPages));
        Assert.Equal(new HeaderFooter(Left: "Sales", Right: "&D"), sheet.PageHeader);
        Assert.Equal(new HeaderFooter(Center: "Page &P of &N"), sheet.PageFooter);
        Assert.Equal(new HeaderFooter(Center: "Cover"), sheet.FirstPageHeader);
        Assert.Equal(new HeaderFooter(), sheet.FirstPageFooter);
        Assert.Null(sheet.EvenPageHeader);
        Assert.Equal((PaperSize.Letter, 100, (PageFit?)null, (HeaderFooter?)null), (workbook["Plain"].PaperSize, workbook["Plain"].PrintScale, workbook["Plain"].FitToPages, workbook["Plain"].PageHeader));
    }

    [Fact]
    public void Unchanged_paper_and_headers_stay_as_the_file_has_them()
    {
        var (original, workbook) = LoadSetup();
        workbook["Setup"]["A1"].Value = 2;

        var (part, _, _) = SaveChecked(original, workbook);

        Assert.Contains("<pageSetup paperSize=\"9\" scale=\"80\" fitToHeight=\"0\" orientation=\"portrait\" r:id=\"rP\"", part);
        Assert.Contains("<headerFooter differentFirst=\"1\" scaleWithDoc=\"0\">", part);
    }

    [Fact]
    public void Changed_paper_scale_fit_and_headers_are_saved()
    {
        var (original, workbook) = LoadSetup();
        var sheet = workbook["Setup"];
        sheet.PaperSize = PaperSize.Legal;
        sheet.FitToPages = null;
        sheet.PrintScale = 120;
        sheet.FirstPageHeader = null;
        sheet.FirstPageFooter = null;
        sheet.EvenPageFooter = new HeaderFooter(Left: "&P");
        var plain = workbook["Plain"];
        plain.FitToPages = new PageFit(1, 1);
        plain.PageFooter = new HeaderFooter(Center: HeaderFooter.Escape("R&D"));

        var (part, _, reloaded) = SaveChecked(original, workbook);

        Assert.Contains("<pageSetup paperSize=\"5\" scale=\"120\" orientation=\"portrait\" r:id=\"rP\"", part);
        Assert.Contains("<headerFooter scaleWithDoc=\"0\" differentOddEven=\"1\"><oddHeader>&amp;LSales&amp;R&amp;D</oddHeader><oddFooter>&amp;CPage &amp;P of &amp;N</oddFooter><evenHeader>&amp;LSales&amp;R&amp;D</evenHeader><evenFooter>&amp;L&amp;P</evenFooter></headerFooter>", part);
        Assert.DoesNotContain("fitToPage", part);
        Assert.Null(reloaded["Setup"].EvenPageHeader);   // written as the odd page's, so none of its own
        Assert.Equal((new PageFit(1, 1), "R&&D"), (reloaded["Plain"].FitToPages, reloaded["Plain"].PageFooter!.Center));
    }

    [Fact]
    public void Removing_every_header_drops_the_element()
    {
        var (original, workbook) = LoadSetup();
        var sheet = workbook["Setup"];
        sheet.PageHeader = sheet.PageFooter = sheet.FirstPageHeader = sheet.FirstPageFooter = null;

        var (part, _, _) = SaveChecked(original, workbook);

        Assert.DoesNotContain("headerFooter", part);
    }

    [Fact]
    public void A_workbook_built_in_code_saves_its_page_setup()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Report");
        sheet["A1"].Value = "x";
        sheet.PaperSize = PaperSize.A4;
        sheet.FitToPages = new PageFit(1, 0);
        sheet.PageHeader = new HeaderFooter(Center: "&\"Arial,Bold\"&14Report");
        sheet.PageFooter = new HeaderFooter(Left: HeaderFooter.SheetName, Right: $"{HeaderFooter.PageNumber} / {HeaderFooter.PageCount}");
        sheet.FirstPageFooter = new HeaderFooter();

        var saved = Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        Assert.Equal(Describe(sheet), Describe(Reload(saved)["Report"]));
    }
}
