using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>Cases a review of 0.13 found broken, kept as tests.</summary>
public class ReviewFindingsTests
{
    private const string PercentStyles = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
          <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs>
        </styleSheet>
        """;

    private static (byte[] Original, Workbook Workbook) Load(string sheet, string? styles = PercentStyles)
    {
        var original = new TestXlsx { Styles = styles }.RawSheet("S", sheet).Build().ToArray();
        return (original, XlsxReader.Load(new MemoryStream(original)));
    }

    [Fact]
    public void A_formatted_row_keeps_its_style_index_when_its_height_changes()
    {
        var (_, workbook) = Load("""<sheetData><row r="2" s="1" customFormat="1"><c r="A2" s="1"><v>0.5</v></c></row></sheetData>""");
        workbook["S"].SetRowHeight(2, 30);
        workbook["S"].GroupRows(2, 2);

        var part = XlsxWriterTests.Part(XlsxWriterTests.Save(workbook), "xl/worksheets/sheet1.xml");

        Assert.Contains("s=\"1\" customFormat=\"1\"", part);
        Assert.Equal("0.00%", workbook["S"].RowNumberFormat(2));
    }

    [Fact]
    public void A_new_cell_in_a_formatted_column_takes_its_number_format()
    {
        var (original, workbook) = Load("""<cols><col min="2" max="2" style="1"/></cols><sheetData/>""");
        var sheet = workbook["S"];

        Assert.Equal("0.00%", sheet["B5"].NumberFormat);
        Assert.Equal("General", sheet["C5"].NumberFormat);
        sheet["B5"].Value = 0.5;
        Assert.Equal("50.00%", sheet["B5"].Text);

        var saved = XlsxWriterTests.Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        Assert.Equal("50.00%", XlsxWriterTests.Reload(saved)["S"]["B5"].Text);
    }

    [Fact]
    public void General_of_its_own_stays_in_a_formatted_row_and_formats_of_lines_are_set_in_code()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet.SetRowNumberFormat(3, "0.0%");
        sheet.SetColumnNumberFormat(2, "#,##0");
        sheet["A3"].Value = 0.25;
        sheet["B4"].Value = 1234;
        sheet["C3"].Value = 0.5;
        sheet["C3"].NumberFormat = "General";

        Assert.Equal(("25.0%", "1,234", "0.5"), (sheet["A3"].Text, sheet["B4"].Text, sheet["C3"].Text));
        Assert.Equal("0.0%", sheet["B3"].NumberFormat);   // the row before the column
        Assert.Throws<ArgumentException>(() => sheet.SetRowNumberFormat(1, "[["));

        var saved = XlsxWriterTests.Save(workbook);
        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        var reloaded = XlsxWriterTests.Reload(saved)["S"];
        Assert.Equal(("25.0%", "1,234", "0.5", "0.0%", "#,##0"), (reloaded["A3"].Text, reloaded["B4"].Text, reloaded["C3"].Text, reloaded.RowNumberFormat(3), reloaded.ColumnNumberFormat(2)));
    }

    [Fact]
    public void Formatted_text_longer_than_excel_allows_loads_as_plain_text()
    {
        var shared = $"""
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><rPr><b/></rPr><t>{new string('a', 20000)}</t></r><r><t>{new string('b', 20000)}</t></r></si></sst>
            """;
        var file = new TestXlsx { SharedStrings = shared }.RawSheet("S", """<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>""").Build().ToArray();

        var workbook = XlsxReader.Load(new MemoryStream(file));
        workbook["S"]["B1"].Value = 1;

        Assert.Equal((40000, (RichText?)null), (workbook["S"]["A1"].Value.AsText().Length, workbook["S"]["A1"].RichText));
        Assert.NotEmpty(XlsxWriterTests.Save(workbook));
    }

    [Fact]
    public void Rows_inserted_in_a_collapsed_group_are_hidden_and_a_deleted_summary_passes_its_mark_on()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet.GroupRows(2, 4);
        sheet.CollapseRowGroup(2);

        sheet.InsertRows(3, 2);
        Assert.Equal([true, true, true, true, true, false], Enumerable.Range(2, 6).Select(sheet.IsRowHidden));
        Assert.True(sheet.IsRowGroupCollapsed(4));

        sheet.DeleteRows(7);   // the summary row
        Assert.True(sheet.IsRowGroupCollapsed(4));
        Assert.Contains(7, sheet.CollapsedRows);
    }

    [Fact]
    public void A_named_style_set_again_in_other_letters_takes_the_new_spelling()
    {
        var workbook = new Workbook();
        workbook.AddSheet("S");
        workbook.SetNamedStyle("Mine", CellStyle.Default with { WrapText = true });
        var reloaded = XlsxWriterTests.Reload(XlsxWriterTests.Save(workbook));

        reloaded.SetNamedStyle("MINE", CellStyle.Default with { WrapText = false, ShrinkToFit = true });
        var saved = XlsxWriterTests.Save(reloaded);
        var styles = XlsxWriterTests.Part(saved, "xl/styles.xml");

        Assert.Equal("MINE", Assert.Single(reloaded.NamedStyles).Key);
        Assert.Contains("<cellStyle name=\"MINE\"", styles);
        Assert.DoesNotContain("xfId=\"1\" applyAlignment", styles.Split("<cellStyleXfs")[1].Split("</cellStyleXfs>")[0]);
        Assert.True(XlsxWriterTests.Reload(saved).NamedStyles["Mine"].ShrinkToFit);
    }

    [Fact]
    public void Fit_to_page_without_a_page_setup_element_fits_to_one_page()
    {
        var (original, workbook) = Load("""<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/>""", styles: null);

        Assert.Equal(new PageFit(1, 1), workbook["S"].FitToPages);
        workbook["S"].FitToPages = null;
        var part = XlsxWriterTests.Part(XlsxWriterTests.Save(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("fitToPage", part);
    }

    [Fact]
    public void A_gradient_angle_must_be_a_number()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => CellGradient.Linear(default, default, double.PositiveInfinity));
        Assert.Throws<ArgumentOutOfRangeException>(() => CellGradient.Linear(default, default) with { Angle = double.NaN });
    }
}
