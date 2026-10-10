using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class SheetViewReaderTests
{
    private static Worksheet Load(string views) =>
        XlsxReader.Load(new TestXlsx().RawSheet("S", $"<sheetViews>{views}</sheetViews><sheetData/>").Build())["S"];

    [Fact]
    public void Frozen_panes_come_from_the_pane_of_the_sheet_view()
    {
        var sheet = Load("<sheetView workbookViewId=\"0\"><pane xSplit=\"1\" ySplit=\"2\" topLeftCell=\"B3\" activePane=\"bottomRight\" state=\"frozen\"/></sheetView>");

        Assert.Equal(2, sheet.FrozenRows);
        Assert.Equal(1, sheet.FrozenColumns);
    }

    [Fact]
    public void Frozen_rows_alone_and_frozenSplit_are_frozen_too()
    {
        Assert.Equal((1, 0), View(Load("<sheetView workbookViewId=\"0\"><pane ySplit=\"1\" state=\"frozen\"/></sheetView>")));
        Assert.Equal((0, 3), View(Load("<sheetView workbookViewId=\"0\"><pane xSplit=\"3\" state=\"frozenSplit\"/></sheetView>")));
    }

    [Fact]
    public void A_split_that_is_not_frozen_freezes_nothing()
    {
        // xSplit of a plain split is a distance in twips, not a column count.
        Assert.Equal((0, 0), View(Load("<sheetView workbookViewId=\"0\"><pane xSplit=\"2400\" ySplit=\"1800\"/></sheetView>")));
    }

    [Fact]
    public void Gridlines_follow_showGridLines()
    {
        Assert.False(Load("<sheetView showGridLines=\"0\" workbookViewId=\"0\"/>").ShowGridlines);
        Assert.True(Load("<sheetView workbookViewId=\"0\"/>").ShowGridlines);
    }

    [Fact]
    public void Only_the_first_sheet_view_counts()
    {
        var sheet = Load("<sheetView workbookViewId=\"0\"><pane ySplit=\"1\" state=\"frozen\"/></sheetView><sheetView showGridLines=\"0\" workbookViewId=\"1\"><pane ySplit=\"5\" state=\"frozen\"/></sheetView>");

        Assert.Equal(1, sheet.FrozenRows);
        Assert.True(sheet.ShowGridlines);
    }

    [Theory]
    [InlineData("<pane xSplit=\"-2\" ySplit=\"abc\" state=\"frozen\"/>")]
    [InlineData("<pane xSplit=\"1.5\" ySplit=\"99999999\" state=\"frozen\"/>")]
    public void Broken_panes_freeze_nothing_and_never_stop_a_file(string pane)
    {
        Assert.Equal((0, 0), View(Load($"<sheetView workbookViewId=\"0\">{pane}</sheetView>")));
    }

    private static (int Rows, int Columns) View(Worksheet sheet) => (sheet.FrozenRows, sheet.FrozenColumns);
}
