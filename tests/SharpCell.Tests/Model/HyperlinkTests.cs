namespace SharpCell.Tests.Model;

/// <summary>A cell can carry a hyperlink; it moves with the cell and is copied with the workbook.</summary>
public class HyperlinkTests
{
    private static Worksheet NewSheet() => new Workbook().AddSheet("S");

    [Fact]
    public void A_link_can_be_set_read_replaced_and_removed()
    {
        var sheet = NewSheet();
        Assert.Null(sheet["B2"].Hyperlink);

        sheet["B2"].Hyperlink = new Hyperlink("https://example.com", "Open");
        Assert.Equal(new Hyperlink("https://example.com", "Open"), sheet["B2"].Hyperlink);

        sheet["B2"].Hyperlink = new Hyperlink("mailto:a@example.com");
        Assert.Equal("mailto:a@example.com", sheet["B2"].Hyperlink!.Target);
        Assert.Null(sheet["B2"].Hyperlink!.Tooltip);

        sheet["B2"].Hyperlink = null;
        Assert.Null(sheet["B2"].Hyperlink);
    }

    [Fact]
    public void Setting_a_link_leaves_the_value_alone()
    {
        var sheet = NewSheet();
        sheet["A1"].Value = "Home";
        sheet["A1"].Hyperlink = new Hyperlink("#Sheet2!A1");

        Assert.Equal(CellValue.Text("Home"), sheet["A1"].Value);
    }

    [Fact]
    public void Links_with_the_same_target_and_tooltip_are_equal()
    {
        Assert.Equal(new Hyperlink("a", "t"), new Hyperlink("a", "t"));
        Assert.Equal(new Hyperlink("a", "t").GetHashCode(), new Hyperlink("a", "t").GetHashCode());
        Assert.NotEqual(new Hyperlink("a", "t"), new Hyperlink("a"));
        Assert.NotEqual(new Hyperlink("a"), new Hyperlink("A"));
        Assert.False(new Hyperlink("a").Equals(null));
    }

    [Theory]
    [InlineData("#Sheet2!A1", true)]
    [InlineData("#'Q1 Sales'!A1", true)]
    [InlineData("#MyName", true)]
    [InlineData("https://example.com/#top", false)]
    [InlineData("report.pdf", false)]
    public void A_target_starting_with_a_hash_is_in_the_workbook(string target, bool inWorkbook)
    {
        Assert.Equal(inWorkbook, new Hyperlink(target).IsInWorkbook);
    }

    [Fact]
    public void An_empty_target_or_a_tooltip_over_255_characters_is_refused()
    {
        Assert.Throws<ArgumentException>(() => new Hyperlink(""));
        Assert.Throws<ArgumentNullException>(() => new Hyperlink(null!));
        Assert.Throws<ArgumentException>(() => new Hyperlink("a", new string('x', 256)));
        Assert.Equal(255, new Hyperlink("a", new string('x', 255)).Tooltip!.Length);
    }

    [Fact]
    public void Inserting_a_row_moves_the_link_with_its_cell()
    {
        var sheet = NewSheet();
        sheet["B2"].Hyperlink = new Hyperlink("a");

        sheet.InsertRows(1);

        Assert.Null(sheet["B2"].Hyperlink);
        Assert.Equal(new Hyperlink("a"), sheet["B3"].Hyperlink);
    }

    [Fact]
    public void Inserting_a_column_moves_the_link_with_its_cell()
    {
        var sheet = NewSheet();
        sheet["B2"].Hyperlink = new Hyperlink("a");

        sheet.InsertColumns(1);

        Assert.Equal(new Hyperlink("a"), sheet["C2"].Hyperlink);
    }

    [Fact]
    public void Deleting_the_row_removes_the_link()
    {
        var sheet = NewSheet();
        sheet["B2"].Hyperlink = new Hyperlink("a");

        sheet.DeleteRows(2);

        Assert.Null(sheet["B2"].Hyperlink);
        Assert.Empty(sheet.Hyperlinks);
    }

    [Fact]
    public void A_clone_keeps_the_links_and_is_independent()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Hyperlink = new Hyperlink("a");

        var clone = workbook.Clone();
        Assert.Equal(new Hyperlink("a"), clone["S"]["A1"].Hyperlink);

        clone["S"]["A1"].Hyperlink = new Hyperlink("b");
        clone["S"]["A2"].Hyperlink = new Hyperlink("c");
        Assert.Equal(new Hyperlink("a"), sheet["A1"].Hyperlink);
        Assert.Null(sheet["A2"].Hyperlink);
    }

    [Fact]
    public void Setting_a_link_from_a_custom_function_is_refused()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        Exception? caught = null;
        workbook.Functions.Add("HOOK", _ =>
        {
            try { sheet["B1"].Hyperlink = new Hyperlink("a"); } catch (Exception ex) { caught = ex; }
            return 5;
        });
        sheet["A1"].Formula = "=HOOK()";
        workbook.Recalculate();

        Assert.IsType<InvalidOperationException>(caught);
        Assert.Null(sheet["B1"].Hyperlink);
    }
}
