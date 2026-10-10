using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>Column widths, row heights, hidden rows and columns, merges, panes and filters are saved.</summary>
public class XlsxWriterSheetTests
{
    private const string Styles = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
          <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
        </styleSheet>
        """;

    private static TestXlsx Full() => new TestXlsx { Styles = Styles }.RawSheet("S", """
        <sheetPr><tabColor rgb="FFFF0000"/></sheetPr>
        <dimension ref="A1:B3"/>
        <sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>
        <sheetFormatPr defaultRowHeight="15"/>
        <cols><col min="1" max="5" width="12" style="1" customWidth="1"/><col min="8" max="8" width="20" customWidth="1"/></cols>
        <sheetData><row r="1" ht="30" customHeight="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
          <row r="3"><c r="A3"><v>3</v></c></row></sheetData>
        <mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells>
        <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        """);

    private static TestXlsx Bare() => new TestXlsx().RawSheet("S", """
        <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
        <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        """);

    public static TheoryData<string, string> Changes => new()
    {
        { "full", "width" }, { "full", "width inside a range" }, { "full", "width removed" }, { "full", "hidden column" },
        { "full", "height" }, { "full", "height of an empty row" }, { "full", "height removed" }, { "full", "hidden row" },
        { "full", "hidden empty row" }, { "full", "default sizes" }, { "full", "merge" }, { "full", "unmerge" },
        { "full", "frozen rows" }, { "full", "frozen columns" }, { "full", "frozen both" }, { "full", "gridlines" },
        { "full", "filter" },
        { "bare", "width" }, { "bare", "hidden column" }, { "bare", "height of an empty row" }, { "bare", "hidden empty row" },
        { "bare", "default sizes" }, { "bare", "merge" }, { "bare", "frozen both" }, { "bare", "gridlines" }, { "bare", "filter" },
    };

    private static void Change(Worksheet s, string change)
    {
        switch (change)
        {
            case "width": s.SetColumnWidth(2, 25); break;
            case "width inside a range": s.SetColumnWidth(3, 30); break;
            case "width removed": s.SetColumnWidth(8, null); break;
            case "hidden column": s.SetColumnHidden(7, true); break;
            case "height": s.SetRowHeight(3, 40); break;
            case "height of an empty row": s.SetRowHeight(6, 25.5); break;
            case "height removed": s.SetRowHeight(1, null); break;
            case "hidden row": s.SetRowHidden(3, true); break;
            case "hidden empty row": s.SetRowHidden(9, true); break;
            case "default sizes": s.DefaultRowHeight = 18; s.DefaultColumnWidth = 10; break;
            case "merge": s.Merge("A5:B6"); break;
            case "unmerge": s.Unmerge("C1:D1"); break;
            case "frozen rows": s.FrozenRows = 1; break;
            case "frozen columns": s.FrozenColumns = 2; break;
            case "frozen both": s.FrozenRows = 2; s.FrozenColumns = 1; break;
            case "gridlines": s.ShowGridlines = false; break;
            case "filter": s.FilterMode = true; break;
        }
    }

    [Theory]
    [MemberData(nameof(Changes))]
    public void A_change_is_saved_and_reads_back(string file, string change)
    {
        var original = (file == "full" ? Full() : Bare()).Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        Change(workbook["S"], change);

        var saved = XlsxWriterTests.Save(workbook);

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var reloaded = XlsxWriterTests.Reload(saved);
        Assert.Equal(Presentation.Of(workbook), Presentation.Of(reloaded));
        Assert.Equal(Structure.Of(workbook), Structure.Of(reloaded));
        Assert.Equal(workbook["S"]["A1"].Value, reloaded["S"]["A1"].Value);
    }

    [Fact]
    public void A_width_inside_a_range_keeps_the_range_style()
    {
        var workbook = XlsxWriterTests.Load(Full());
        workbook["S"].SetColumnWidth(3, 30);

        var sheet = XlsxWriterTests.Part(XlsxWriterTests.Save(workbook), "xl/worksheets/sheet1.xml");

        Assert.Contains("<col min=\"1\" max=\"2\" width=\"12\" style=\"1\" customWidth=\"1\"", sheet);
        Assert.Contains("<col min=\"3\" max=\"3\" width=\"30\" style=\"1\" customWidth=\"1\"", sheet);
        Assert.Contains("<col min=\"4\" max=\"5\" width=\"12\" style=\"1\" customWidth=\"1\"", sheet);
    }

    [Fact]
    public void Freezing_writes_a_frozen_pane_and_unfreezing_removes_it()
    {
        var workbook = XlsxWriterTests.Load(Full());
        workbook["S"].FrozenRows = 1;
        workbook["S"].FrozenColumns = 2;
        var frozen = XlsxWriterTests.Save(workbook);

        Assert.Contains("<pane xSplit=\"2\" ySplit=\"1\" topLeftCell=\"C2\" activePane=\"bottomRight\" state=\"frozen\"", XlsxWriterTests.Part(frozen, "xl/worksheets/sheet1.xml"));

        var again = XlsxWriterTests.Reload(frozen);
        again.Recalculate();
        again["S"].FrozenRows = 0;
        again["S"].FrozenColumns = 0;
        var thawed = XlsxWriterTests.Part(XlsxWriterTests.Save(again), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("<pane", thawed);
        Assert.DoesNotContain("pane=", thawed);
    }

    [Fact]
    public void A_hidden_empty_row_gets_a_row_element()
    {
        var workbook = XlsxWriterTests.Load(Bare());
        workbook["S"].SetRowHidden(4, true);

        Assert.Contains("<row r=\"4\" hidden=\"1\"", XlsxWriterTests.Part(XlsxWriterTests.Save(workbook), "xl/worksheets/sheet1.xml"));
    }
}
