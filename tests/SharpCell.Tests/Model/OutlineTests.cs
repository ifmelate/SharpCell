namespace SharpCell.Tests.Model;

/// <summary>Grouped rows and columns: levels, collapsing and expanding, and how they follow edits.</summary>
public class OutlineTests
{
    [Fact]
    public void Grouping_puts_rows_one_level_deeper_and_ungrouping_takes_them_out()
    {
        var sheet = new Workbook().AddSheet("S");

        sheet.GroupRows(2, 9);
        sheet.GroupRows(3, 5);
        sheet.GroupColumns(2, 3);

        Assert.Equal([0, 1, 2, 2, 2, 1, 1, 1, 1, 0], Enumerable.Range(1, 10).Select(sheet.RowOutlineLevel));
        Assert.Equal((0, 1, 1, 0), (sheet.ColumnOutlineLevel(1), sheet.ColumnOutlineLevel(2), sheet.ColumnOutlineLevel(3), sheet.ColumnOutlineLevel(4)));
        sheet.UngroupRows(1, 4);
        Assert.Equal([0, 0, 1, 1, 2, 1], Enumerable.Range(1, 6).Select(sheet.RowOutlineLevel));
    }

    [Fact]
    public void Levels_stop_at_seven()
    {
        var sheet = new Workbook().AddSheet("S");
        for (var i = 0; i < 7; i++)
            sheet.GroupRows(2, 3);

        Assert.Throws<InvalidOperationException>(() => sheet.GroupRows(1, 2));
        Assert.Equal(0, sheet.RowOutlineLevel(1));   // nothing changed
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetRowOutlineLevel(1, 8));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.GroupRows(5, 4));
    }

    [Fact]
    public void Collapsing_hides_the_group_and_marks_the_summary_row_below()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet.GroupRows(2, 4);

        sheet.CollapseRowGroup(3);

        Assert.Equal([false, true, true, true, false], Enumerable.Range(1, 5).Select(sheet.IsRowHidden));
        Assert.True(sheet.IsRowGroupCollapsed(2));
        Assert.Contains(5, sheet.CollapsedRows);
        sheet.ExpandRowGroup(2);
        Assert.DoesNotContain(true, Enumerable.Range(1, 5).Select(sheet.IsRowHidden));
        Assert.False(sheet.IsRowGroupCollapsed(4));
        Assert.Throws<ArgumentException>(() => sheet.CollapseRowGroup(7));
    }

    [Fact]
    public void A_summary_above_is_marked_instead_when_summaries_are_not_below()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet.OutlineSummaryBelow = false;
        sheet.GroupRows(3, 4);

        sheet.CollapseRowGroup(3);

        Assert.Contains(2, sheet.CollapsedRows);
        Assert.False(sheet.IsRowHidden(2));
    }

    [Fact]
    public void Expanding_an_outer_group_keeps_an_inner_collapsed_group_hidden()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet.GroupRows(2, 8);
        sheet.GroupRows(3, 5);
        sheet.CollapseRowGroup(4);
        sheet.CollapseRowGroup(2);

        sheet.ExpandRowGroup(2);

        Assert.Equal([false, false, true, true, true, false, false, false], Enumerable.Range(1, 8).Select(sheet.IsRowHidden));
        Assert.True(sheet.IsRowGroupCollapsed(3));
        Assert.False(sheet.IsRowGroupCollapsed(7));
    }

    [Fact]
    public void Columns_collapse_left_of_their_summary_column()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet.GroupColumns(2, 3);

        sheet.CollapseColumnGroup(2);

        Assert.Equal((false, true, true, false), (sheet.IsColumnHidden(1), sheet.IsColumnHidden(2), sheet.IsColumnHidden(3), sheet.IsColumnHidden(4)));
        Assert.True(sheet.IsColumnGroupCollapsed(3));
        sheet.ExpandColumnGroup(3);
        Assert.False(sheet.IsColumnHidden(2));
    }

    [Fact]
    public void Rows_inserted_inside_a_group_join_it_and_levels_follow_edits()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet.GroupRows(3, 6);
        sheet.CollapseRowGroup(3);

        sheet.InsertRows(4, 2);
        Assert.Equal([0, 0, 1, 1, 1, 1, 1, 1, 0], Enumerable.Range(1, 9).Select(sheet.RowOutlineLevel));
        Assert.Contains(9, sheet.CollapsedRows);

        sheet.InsertRows(3);   // at the edge: not inside
        Assert.Equal(0, sheet.RowOutlineLevel(3));
        sheet.DeleteRows(4, 3);
        Assert.Equal([0, 0, 0, 1, 1, 1, 0], Enumerable.Range(1, 7).Select(sheet.RowOutlineLevel));
    }

    [Fact]
    public void Subtotal_skips_collapsed_rows_only_with_the_hundreds_codes()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        for (var row = 1; row <= 4; row++)
            sheet[$"A{row}"].Value = row;
        sheet["B1"].Formula = "=SUBTOTAL(9,A1:A4)";
        sheet["B2"].Formula = "=SUBTOTAL(109,A1:A4)";
        sheet.GroupRows(2, 3);

        sheet.CollapseRowGroup(2);
        workbook.Recalculate();

        Assert.Equal((10.0, 5.0), (sheet["B1"].Value.AsNumber(), sheet["B2"].Value.AsNumber()));
    }

    [Fact]
    public void A_clone_has_the_same_outline()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet.GroupRows(2, 3);
        sheet.CollapseRowGroup(2);
        sheet.OutlineSummaryRight = false;

        var clone = workbook.Clone()["S"];

        Assert.Equal((1, true, false), (clone.RowOutlineLevel(2), clone.IsRowGroupCollapsed(3), clone.OutlineSummaryRight));
    }
}
