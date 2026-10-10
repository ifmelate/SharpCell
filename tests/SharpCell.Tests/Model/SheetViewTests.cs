namespace SharpCell.Tests.Model;

public class SheetViewTests
{
    private static Worksheet Sheet() => new Workbook().AddSheet("S");

    [Fact]
    public void A_new_sheet_has_no_frozen_panes_and_shows_gridlines()
    {
        var sheet = Sheet();

        Assert.Equal(0, sheet.FrozenRows);
        Assert.Equal(0, sheet.FrozenColumns);
        Assert.True(sheet.ShowGridlines);
    }

    [Fact]
    public void Frozen_panes_and_gridlines_are_kept()
    {
        var sheet = Sheet();
        sheet.FrozenRows = 2;
        sheet.FrozenColumns = 1;
        sheet.ShowGridlines = false;

        Assert.Equal((2, 1, false), (sheet.FrozenRows, sheet.FrozenColumns, sheet.ShowGridlines));
    }

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(0, -1)]
    [InlineData(1_048_576, 0)]
    [InlineData(0, 16_384)]
    public void Frozen_panes_must_leave_part_of_the_sheet_to_scroll(int rows, int columns)
    {
        var sheet = Sheet();

        Assert.Throws<ArgumentOutOfRangeException>(() =>
        {
            sheet.FrozenRows = rows;
            sheet.FrozenColumns = columns;
        });
    }

    [Fact]
    public void A_clone_has_the_same_view()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet.FrozenRows = 3;
        sheet.FrozenColumns = 2;
        sheet.ShowGridlines = false;

        var copy = workbook.Clone()["S"];

        Assert.Equal((3, 2, false), (copy.FrozenRows, copy.FrozenColumns, copy.ShowGridlines));
    }
}
