namespace SharpCell.Tests.Model;

/// <summary>A sheet's filter hides the rows its criteria do not show, as Excel's Data &gt; Filter does.</summary>
public class AutoFilterTests
{
    // A: region, B: units, C: date. Rows 2 to 7.
    private static (Workbook W, Worksheet S) Book()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = "Region";
        sheet["B1"].Value = "Units";
        sheet["C1"].Value = "Date";
        object[][] rows =
        [
            ["North", 1250.0, new DateTime(2026, 3, 5)],
            ["south", 40.0, new DateTime(2026, 3, 20)],
            ["East", 300.0, new DateTime(2026, 4, 1)],
            ["", 75.0, new DateTime(2025, 3, 5)],
            ["North-east", 300.0, new DateTime(2026, 10, 17)],
            ["West", 5.0, new DateTime(2026, 10, 1)],
        ];
        for (var i = 0; i < rows.Length; i++)
        {
            if ((string)rows[i][0] != "")
                sheet[i + 2, 1].Value = (string)rows[i][0];
            sheet[i + 2, 2].Value = (double)rows[i][1];
            sheet[i + 2, 3].Value = CellValue.DateTime((DateTime)rows[i][2], DateSystem.Date1900);
        }

        for (var row = 2; row <= 7; row++)
        {
            sheet[row, 2].NumberFormat = "#,##0.00";
            sheet[row, 3].NumberFormat = "yyyy-mm-dd";
        }
        workbook.Recalculate();
        return (workbook, sheet);
    }

    private static int[] Shown(Worksheet sheet) => [.. Enumerable.Range(2, 6).Where(row => !sheet.IsRowHidden(row))];

    [Fact]
    public void Values_are_matched_as_the_cells_show_them_ignoring_case()
    {
        var (_, s) = Book();
        var filter = s.SetAutoFilter("A1:C7");

        filter.Filter(2, ColumnFilter.OneOf("1,250.00", "300.00"));
        Assert.Equal([2, 4, 6], Shown(s));
        Assert.True(s.FilterMode);

        filter.Filter(2, ColumnFilter.OneOf("40"));
        Assert.Empty(Shown(s));

        filter.Filter(1, ColumnFilter.OneOf("SOUTH", "north") with { Blanks = true });
        filter.Clear(2);
        Assert.Equal([2, 3, 5], Shown(s));
    }

    [Fact]
    public void Blank_shows_only_empty_cells()
    {
        var (_, s) = Book();

        s.SetAutoFilter("A1:C7").Filter(1, ColumnFilter.Blank());

        Assert.Equal([5], Shown(s));
    }

    [Fact]
    public void Dates_are_matched_by_year_month_or_day()
    {
        var (_, s) = Book();
        var filter = s.SetAutoFilter("A1:C7");

        filter.Filter(3, ColumnFilter.OneOf(new FilterDate(DateGrouping.Month, 2026, 3), new FilterDate(DateGrouping.Day, 2026, 10, 17)));
        Assert.Equal([2, 3, 6], Shown(s));

        filter.Filter(3, ColumnFilter.OneOf(new FilterDate(DateGrouping.Year, 2025)));
        Assert.Equal([5], Shown(s));
    }

    [Fact]
    public void Custom_conditions_compare_numbers_as_numbers_and_text_with_wildcards()
    {
        var (_, s) = Book();
        var filter = s.SetAutoFilter("A1:C7");

        filter.Filter(2, ColumnFilter.Custom(FilterOperator.GreaterThanOrEqual, "300"));
        Assert.Equal([2, 4, 6], Shown(s));

        filter.Filter(2, ColumnFilter.Custom(new FilterCondition(FilterOperator.LessThan, "50"), new FilterCondition(FilterOperator.GreaterThan, "1000"), matchAll: false));
        Assert.Equal([2, 3, 7], Shown(s));

        filter.Clear(2);
        filter.Filter(1, ColumnFilter.Custom(FilterOperator.Equal, "n*"));
        Assert.Equal([2, 6], Shown(s));
        filter.Filter(1, ColumnFilter.BeginsWith("North-"));
        Assert.Equal([6], Shown(s));
        filter.Filter(1, ColumnFilter.Contains("ST"));
        Assert.Equal([4, 6, 7], Shown(s));
        filter.Filter(1, ColumnFilter.EndsWith("h"));
        Assert.Equal([2, 3], Shown(s));
        // Not containing a text shows blanks too.
        filter.Filter(1, ColumnFilter.DoesNotContain("st"));
        Assert.Equal([2, 3, 5], Shown(s));
        filter.Filter(1, ColumnFilter.Custom(FilterOperator.GreaterThan, "o"));
        Assert.Equal([3, 7], Shown(s));
    }

    [Fact]
    public void A_wildcard_written_into_a_text_filter_is_matched_literally()
    {
        var (w, s) = Book();
        s["A2"].Value = "a*b";
        s["A3"].Value = "axb";
        w.Recalculate();

        s.SetAutoFilter("A1:C7").Filter(1, ColumnFilter.Contains("*"));

        Assert.Equal([2], Shown(s));
    }

    [Fact]
    public void Top_and_bottom_show_the_highest_and_lowest_numbers_with_ties()
    {
        var (_, s) = Book();
        var filter = s.SetAutoFilter("A1:C7");

        filter.Filter(2, ColumnFilter.Top(2));
        Assert.Equal([2, 4, 6], Shown(s));
        filter.Filter(2, ColumnFilter.Bottom(1));
        Assert.Equal([7], Shown(s));
        filter.Filter(2, ColumnFilter.Top(50, percent: true));
        Assert.Equal([2, 4, 6], Shown(s));
    }

    [Fact]
    public void Above_and_below_average_compare_with_the_columns_average()
    {
        var (_, s) = Book();
        var filter = s.SetAutoFilter("A1:C7");

        // The average of 1250, 40, 300, 75, 300 and 5 is 328.33.
        filter.Filter(2, ColumnFilter.Of(DynamicFilter.AboveAverage));
        Assert.Equal([2], Shown(s));
        filter.Filter(2, ColumnFilter.Of(DynamicFilter.BelowAverage));
        Assert.Equal([3, 4, 5, 6, 7], Shown(s));
    }

    [Fact]
    public void Date_periods_count_from_today()
    {
        var (w, s) = Book();
        w.Clock = new FixedClock(new DateTime(2026, 10, 17, 9, 0, 0));
        var filter = s.SetAutoFilter("A1:C7");

        filter.Filter(3, ColumnFilter.Of(DynamicFilter.ThisMonth));
        Assert.Equal([6, 7], Shown(s));
        filter.Filter(3, ColumnFilter.Of(DynamicFilter.Today));
        Assert.Equal([6], Shown(s));
        filter.Filter(3, ColumnFilter.Of(DynamicFilter.March));
        Assert.Equal([2, 3, 5], Shown(s));
        filter.Filter(3, ColumnFilter.Of(DynamicFilter.Quarter2));
        Assert.Equal([4], Shown(s));
        filter.Filter(3, ColumnFilter.Of(DynamicFilter.LastYear));
        Assert.Equal([5], Shown(s));
        // Saturday 17 October: this week runs from Sunday the 11th.
        filter.Filter(3, ColumnFilter.Of(DynamicFilter.ThisWeek));
        Assert.Equal([6], Shown(s));
    }

    [Fact]
    public void Columns_combine_and_reapplying_follows_changed_data()
    {
        var (w, s) = Book();
        var filter = s.SetAutoFilter("A1:C7");
        filter.Filter(1, ColumnFilter.BeginsWith("North"));
        filter.Filter(2, ColumnFilter.Custom(FilterOperator.LessThan, "1000"));
        Assert.Equal([6], Shown(s));

        s["B2"].Value = 10;
        s["A3"].Value = "Northwest";
        w.Recalculate();
        // Nothing moves until the filter is applied again, as in Excel.
        Assert.Equal([6], Shown(s));

        filter.Apply();
        Assert.Equal([2, 3, 6], Shown(s));
    }

    [Fact]
    public void Formulas_out_of_date_are_calculated_before_filtering()
    {
        var (_, s) = Book();
        s["D1"].Value = "Double";
        for (var row = 2; row <= 7; row++)
            s[row, 4].Formula = $"=B{row}*2";

        s.SetAutoFilter("A1:D7").Filter(4, ColumnFilter.Custom(FilterOperator.GreaterThan, "500"));

        Assert.Equal([2, 4, 6], Shown(s));
    }

    [Fact]
    public void Clearing_and_removing_the_filter_shows_the_rows()
    {
        var (_, s) = Book();
        var filter = s.SetAutoFilter("A1:C7");
        filter.Filter(2, ColumnFilter.Top(1));

        filter.Clear();
        Assert.Equal([2, 3, 4, 5, 6, 7], Shown(s));
        Assert.False(s.FilterMode);
        Assert.Same(filter, s.AutoFilter);

        filter.Filter(2, ColumnFilter.Top(1));
        s.RemoveAutoFilter();
        Assert.Null(s.AutoFilter);
        Assert.Equal([2, 3, 4, 5, 6, 7], Shown(s));
        Assert.Throws<InvalidOperationException>(() => filter.Apply());
    }

    [Fact]
    public void A_filter_over_whole_columns_ends_at_the_last_row_used()
    {
        var (_, s) = Book();

        var filter = s.SetAutoFilter("A:C");

        Assert.Equal("A1:C7", filter.Range.Address);
    }

    [Fact]
    public void A_filter_cannot_overlap_a_table_or_be_one_row()
    {
        var (_, s) = Book();
        s.AddTable("T", "E1:F3", hasHeaderRow: false);

        Assert.Throws<ArgumentException>(() => s.SetAutoFilter("D1:E5"));
        Assert.Throws<ArgumentException>(() => s.SetAutoFilter("A1:C1"));
        Assert.Throws<ArgumentOutOfRangeException>(() => s.SetAutoFilter("A1:C7").Filter(4, ColumnFilter.Blank()));
    }

    [Fact]
    public void The_filter_moves_with_inserted_rows_and_columns()
    {
        var (_, s) = Book();
        var filter = s.SetAutoFilter("A1:C7");
        filter.Filter(2, ColumnFilter.Top(1));

        s.InsertRows(4);
        s.InsertColumns(1);

        Assert.Equal("B1:D8", filter.Range.Address);
        Assert.Equal([3], filter.Columns.Keys);
        Assert.Equal(ColumnFilterType.Top, filter.Columns[3].Type);

        s.DeleteRows(1, 8);
        Assert.Null(s.AutoFilter);
    }

    [Fact]
    public void A_clone_has_its_own_filter()
    {
        var (w, s) = Book();
        s.SetAutoFilter("A1:C7").Filter(2, ColumnFilter.Top(1));

        var copy = w.Clone()["S"];
        copy.AutoFilter!.Clear();

        Assert.Equal([2], Shown(s));
        Assert.Equal([2, 3, 4, 5, 6, 7], Shown(copy));
    }

    private sealed class FixedClock(DateTime local) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => new(local, TimeSpan.Zero);

        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }
}
