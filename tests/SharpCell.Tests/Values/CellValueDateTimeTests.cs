namespace SharpCell.Tests.Values;

public class CellValueDateTimeTests
{
    [Theory]
    [InlineData(1, DateSystem.Date1900, "1900-01-01T00:00:00")]
    [InlineData(59, DateSystem.Date1900, "1900-02-28T00:00:00")]
    [InlineData(61, DateSystem.Date1900, "1900-03-01T00:00:00")]
    [InlineData(46096.5, DateSystem.Date1900, "2026-03-15T12:00:00")]
    [InlineData(0.25, DateSystem.Date1900, "1899-12-31T06:00:00")]
    [InlineData(2958465, DateSystem.Date1900, "9999-12-31T00:00:00")]
    [InlineData(0, DateSystem.Date1904, "1904-01-01T00:00:00")]
    [InlineData(44634.75, DateSystem.Date1904, "2026-03-15T18:00:00")]
    public void AsDateTime_reads_a_serial_number(double serial, DateSystem system, string expected)
    {
        var value = CellValue.Number(serial).AsDateTime(system);

        Assert.Equal(DateTime.Parse(expected, System.Globalization.CultureInfo.InvariantCulture), value);
        Assert.Equal(DateTimeKind.Unspecified, value.Kind);
    }

    [Fact]
    public void AsDateTime_rounds_the_time_to_milliseconds()
    {
        // 1 second short of noon, plus a sliver of floating point error.
        var serial = 46096 + (12 * 3600 - 1) / 86400.0 + 1e-12;
        Assert.Equal(new DateTime(2026, 3, 15, 11, 59, 59), CellValue.Number(serial).AsDateTime(DateSystem.Date1900));
    }

    [Theory]
    [InlineData(-1, DateSystem.Date1900)]
    [InlineData(60, DateSystem.Date1900)]
    [InlineData(60.5, DateSystem.Date1900)]
    [InlineData(2958466, DateSystem.Date1900)]
    [InlineData(2957004, DateSystem.Date1904)]
    public void AsDateTime_refuses_numbers_that_are_no_date(double serial, DateSystem system)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => CellValue.Number(serial).AsDateTime(system));
    }

    [Fact]
    public void AsDateTime_of_a_value_that_is_not_a_number_throws_like_other_As_methods()
    {
        Assert.Throws<InvalidOperationException>(() => CellValue.Text("2026-03-15").AsDateTime(DateSystem.Date1900));
        Assert.Throws<InvalidOperationException>(() => CellValue.Empty.AsDateTime(DateSystem.Date1900));
    }

    [Theory]
    [InlineData("1900-01-01T00:00:00", DateSystem.Date1900, 1)]
    [InlineData("1900-03-01T00:00:00", DateSystem.Date1900, 61)]
    [InlineData("2026-03-15T12:00:00", DateSystem.Date1900, 46096.5)]
    [InlineData("1904-01-01T00:00:00", DateSystem.Date1904, 0)]
    public void DateTime_makes_a_serial_number(string date, DateSystem system, double expected)
    {
        var value = CellValue.DateTime(DateTime.Parse(date, System.Globalization.CultureInfo.InvariantCulture), system);
        Assert.Equal(CellValue.Number(expected), value);
    }

    [Fact]
    public void DateTime_refuses_dates_before_the_date_system()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => CellValue.DateTime(new DateTime(1899, 12, 30), DateSystem.Date1900));
        Assert.Throws<ArgumentOutOfRangeException>(() => CellValue.DateTime(new DateTime(1903, 12, 31), DateSystem.Date1904));
    }

    [Fact]
    public void A_date_round_trips_through_a_cell_and_a_formula()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = CellValue.DateTime(new DateTime(2026, 3, 15), workbook.DateSystem);
        sheet["A2"].Formula = "=A1+1";
        workbook.Recalculate();

        Assert.Equal(new DateTime(2026, 3, 16), sheet["A2"].Value.AsDateTime(workbook.DateSystem));
        Assert.Equal(CellValue.Number(15), workbook.Evaluate("=DAY(S!A1)"));
    }
}
