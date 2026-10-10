namespace SharpCell.Tests.Model;

public class SheetGeometryTests
{
    private static Worksheet Sheet() => new Workbook().AddSheet("S");

    [Fact]
    public void Sizes_are_unset_until_given()
    {
        var sheet = Sheet();

        Assert.Null(sheet.DefaultColumnWidth);
        Assert.Null(sheet.DefaultRowHeight);
        Assert.Null(sheet.ColumnWidth(1));
        Assert.Null(sheet.RowHeight(1));
        Assert.False(sheet.IsColumnHidden(1));
    }

    [Fact]
    public void Widths_heights_and_hidden_columns_are_kept_and_null_forgets_them()
    {
        var sheet = Sheet();
        sheet.DefaultColumnWidth = 9.140625;
        sheet.DefaultRowHeight = 15;
        sheet.SetColumnWidth(2, 20.5);
        sheet.SetRowHeight(3, 30);
        sheet.SetColumnHidden(4, true);

        Assert.Equal(9.140625, sheet.DefaultColumnWidth);
        Assert.Equal(15, sheet.DefaultRowHeight);
        Assert.Equal(20.5, sheet.ColumnWidth(2));
        Assert.Equal(30, sheet.RowHeight(3));
        Assert.True(sheet.IsColumnHidden(4));

        sheet.SetColumnWidth(2, null);
        sheet.SetRowHeight(3, null);
        sheet.SetColumnHidden(4, false);
        Assert.Null(sheet.ColumnWidth(2));
        Assert.Null(sheet.RowHeight(3));
        Assert.False(sheet.IsColumnHidden(4));
    }

    [Fact]
    public void Sizes_outside_Excels_limits_are_refused()
    {
        var sheet = Sheet();

        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(1, 256));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(1, -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetRowHeight(1, 410));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.DefaultRowHeight = double.NaN);
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.ColumnWidth(16385));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.RowHeight(0));
    }

    [Fact]
    public void Hidden_columns_do_not_change_results()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = 1;
        sheet["B1"].Value = 2;
        sheet["C1"].Formula = "=SUBTOTAL(109,A1:B1)";
        sheet.SetColumnHidden(2, true);
        workbook.Recalculate();

        // Excel's SUBTOTAL skips hidden rows, never hidden columns.
        Assert.Equal(CellValue.Number(3), sheet["C1"].Value);
    }

    [Fact]
    public void Merged_areas_are_kept_in_order_and_can_be_undone()
    {
        var sheet = Sheet();
        sheet.Merge("A1:B2");
        sheet.Merge("D4:D6");

        Assert.Equal(["A1:B2", "D4:D6"], sheet.MergedAreas.Select(a => a.Address));

        sheet.Unmerge("A1:B2");
        Assert.Equal(["D4:D6"], sheet.MergedAreas.Select(a => a.Address));
    }

    [Fact]
    public void Merging_keeps_the_values_of_the_cells()
    {
        var sheet = Sheet();
        sheet["B1"].Value = 5;

        sheet.Merge("A1:B1");

        Assert.Equal(CellValue.Number(5), sheet["B1"].Value);
    }

    [Theory]
    [InlineData("C3")]
    [InlineData("B2:C3")]
    [InlineData("A1:A1")]
    public void A_merge_of_one_cell_or_over_another_is_refused(string address)
    {
        var sheet = Sheet();
        sheet.Merge("A1:B2");

        Assert.Throws<ArgumentException>(() => sheet.Merge(address));
        Assert.Single(sheet.MergedAreas);
    }

    [Fact]
    public void Only_an_existing_merge_can_be_undone()
    {
        var sheet = Sheet();
        sheet.Merge("A1:B2");

        Assert.Throws<ArgumentException>(() => sheet.Unmerge("A1:B1"));
    }

    [Fact]
    public void A_clone_has_the_same_geometry()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet.DefaultColumnWidth = 10;
        sheet.SetColumnWidth(3, 12);
        sheet.SetRowHeight(2, 40);
        sheet.SetColumnHidden(5, true);
        sheet.Merge("A1:C1");

        var copy = workbook.Clone()["S"];

        Assert.Equal(10, copy.DefaultColumnWidth);
        Assert.Equal(12, copy.ColumnWidth(3));
        Assert.Equal(40, copy.RowHeight(2));
        Assert.True(copy.IsColumnHidden(5));
        Assert.Equal("A1:C1", Assert.Single(copy.MergedAreas).Address);
        Assert.Same(copy, copy.MergedAreas[0].Worksheet);
    }
}
