using System.Linq;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class HyperlinkReaderTests
{
    private const string Hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    private static string External(string id, string target) =>
        $"<Relationship Id=\"{id}\" Type=\"{Hyperlink}\" Target=\"{target}\" TargetMode=\"External\"/>";

    private static Worksheet Load(string hyperlinks, string relationships = "") =>
        XlsxReader.Load(new TestXlsx()
            .RawSheet("Sheet1", $"<sheetData/><hyperlinks>{hyperlinks}</hyperlinks>")
            .SheetRelationships("Sheet1", relationships)
            .Build())["Sheet1"];

    [Fact]
    public void An_external_link_keeps_its_address_and_tooltip()
    {
        var sheet = Load("<hyperlink ref=\"A1\" r:id=\"rId1\" tooltip=\"Open it\" display=\"Site\"/>", External("rId1", "https://example.com/a?b=1&amp;c=2"));

        var link = Assert.Single(sheet.Hyperlinks);
        Assert.Equal(new CellAddress(1, 1), link.Key);
        Assert.Equal("https://example.com/a?b=1&c=2", link.Value.Target);
        Assert.Equal("Open it", link.Value.Tooltip);
    }

    [Fact]
    public void Mailto_and_relative_file_targets_come_through_raw()
    {
        var sheet = Load(
            "<hyperlink ref=\"A1\" r:id=\"rId1\"/><hyperlink ref=\"A2\" r:id=\"rId2\"/>",
            External("rId1", "mailto:a@example.com?subject=Hi") + External("rId2", "..\\docs\\a.pdf"));

        Assert.Equal("mailto:a@example.com?subject=Hi", sheet.Hyperlinks[new CellAddress(1, 1)].Target);
        Assert.Equal("..\\docs\\a.pdf", sheet.Hyperlinks[new CellAddress(2, 1)].Target);
    }

    [Fact]
    public void A_location_after_an_external_address_becomes_its_fragment()
    {
        var sheet = Load("<hyperlink ref=\"A1\" r:id=\"rId1\" location=\"part\"/>", External("rId1", "https://example.com/page"));

        Assert.Equal("https://example.com/page#part", sheet.Hyperlinks[new CellAddress(1, 1)].Target);
    }

    [Fact]
    public void Escaped_characters_in_a_tooltip_are_decoded()
    {
        var sheet = Load("<hyperlink ref=\"A1\" location=\"B2\" tooltip=\"a_x0001_b\"/><hyperlink ref=\"A2\" location=\"B2\" tooltip=\"_x005F_x0041_\"/>");

        Assert.Equal("a\u0001b", sheet.Hyperlinks[new CellAddress(1, 1)].Tooltip);
        Assert.Equal("_x0041_", sheet.Hyperlinks[new CellAddress(2, 1)].Tooltip);
    }

    [Fact]
    public void A_link_inside_the_workbook_has_only_a_location()
    {
        var sheet = Load("<hyperlink ref=\"B2\" location=\"'Q1 Sales'!A1\"/>");

        var link = sheet.Hyperlinks[new CellAddress(2, 2)];
        Assert.Equal("#'Q1 Sales'!A1", link.Target);
        Assert.True(link.IsInWorkbook);
    }

    [Fact]
    public void A_range_gives_every_cell_the_same_link()
    {
        var sheet = Load("<hyperlink ref=\"A1:B2\" r:id=\"rId1\"/>", External("rId1", "https://example.com"));

        Assert.Equal(4, sheet.Hyperlinks.Count);
        var first = sheet.Hyperlinks[new CellAddress(1, 1)];
        foreach (var link in sheet.Hyperlinks.Values)
            Assert.Same(first, link);
    }

    [Fact]
    public void Links_that_cannot_be_read_are_skipped()
    {
        var sheet = Load(
            "<hyperlink ref=\"A:A\" r:id=\"rId1\"/>"
            + "<hyperlink ref=\"A1:A10001\" r:id=\"rId1\"/>"
            + "<hyperlink ref=\"C1\" r:id=\"rId9\"/>"
            + "<hyperlink ref=\"D1\"/>"
            + "<hyperlink ref=\"E1\" location=\"\"/>"
            + "<hyperlink ref=\"nonsense\" location=\"A1\"/>"
            + "<hyperlink ref=\"F1\" r:id=\"rId1\"/>",
            External("rId1", "https://example.com"));

        var link = Assert.Single(sheet.Hyperlinks);
        Assert.Equal(new CellAddress(1, 6), link.Key);
    }

    [Fact]
    public void An_overlong_tooltip_is_dropped_but_the_link_is_kept()
    {
        var sheet = Load($"<hyperlink ref=\"A1\" location=\"B2\" tooltip=\"{new string('x', 256)}\"/>");

        var link = sheet.Hyperlinks[new CellAddress(1, 1)];
        Assert.Equal("#B2", link.Target);
        Assert.Null(link.Tooltip);
    }

    [Fact]
    public void Links_stop_being_modelled_once_a_sheet_budget_is_spent()
    {
        // Each A1:CV100 covers 10,000 cells; the budget is 100,000, so the eleventh is kept as the file has it.
        var elements = string.Concat(Enumerable.Range(0, 10).Select(_ => "<hyperlink ref=\"A1:CV100\" location=\"B2\"/>"))
            + "<hyperlink ref=\"E200\" location=\"B2\"/>";
        var sheet = Load(elements + "<hyperlink ref=\"A1:CV100\" location=\"Z9\"/>");

        Assert.Equal(10_000, sheet.Hyperlinks.Count);
        Assert.All(sheet.Hyperlinks.Values, link => Assert.Equal("#B2", link.Target));
        Assert.False(sheet.Hyperlinks.ContainsKey(new CellAddress(200, 5)));
    }

    [Fact]
    public void Where_modelled_links_overlap_the_later_one_wins()
    {
        var sheet = Load("<hyperlink ref=\"A1:A2\" location=\"B1\"/><hyperlink ref=\"A2\" location=\"C1\"/>");

        Assert.Equal("#B1", sheet.Hyperlinks[new CellAddress(1, 1)].Target);
        Assert.Equal("#C1", sheet.Hyperlinks[new CellAddress(2, 1)].Target);
    }
}
