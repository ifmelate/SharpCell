using System.Globalization;
using SharpCell;

namespace SharpCell.Tests.Values;

public class DateTextTests
{
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;
    private static readonly CultureInfo EnGb = CultureInfo.GetCultureInfo("en-GB");
    private static readonly CultureInfo RuRu = CultureInfo.GetCultureInfo("ru-RU");

    private static double? Parse(string text, CultureInfo? culture = null, DateSystem system = DateSystem.Date1900, int? year = null) =>
        DateText.TryParse(text, culture ?? Invariant, system, year, out var serial) ? serial : null;

    [Theory]
    [InlineData("12:00", 0.5)]
    [InlineData("12:00:30", 43230.0 / 86400)]
    [InlineData("1:30 PM", 13.5 / 24)]
    [InlineData("1:30 pm", 13.5 / 24)]
    [InlineData("4 pm", 16.0 / 24)]
    [InlineData("4:5 pm", 57900.0 / 86400)]
    [InlineData("4:05:20 pm", 57920.0 / 86400)]
    [InlineData("12:00 AM", 0)]
    [InlineData("12:00 PM", 0.5)]
    [InlineData("0:00", 0)]
    [InlineData("24:00", 1)]
    [InlineData("25:00", 25.0 / 24)]
    [InlineData("23:59:60", 1)]
    [InlineData("4 : 35", 16500.0 / 86400)]
    [InlineData("4:35:00.5", 16500.5 / 86400)]
    [InlineData("12:30.5", 750.5 / 86400)]
    public void Times(string text, double expected) => Assert.Equal(expected, Parse(text)!.Value, 12);

    [Theory]
    [InlineData("2020-01-01", 43831)]
    [InlineData("2020/1/1", 43831)]
    [InlineData("2026 - 01 - 01", 46023)]
    [InlineData("1/2/2020", 43832)]
    [InlineData("1-2-2020", 43832)]
    [InlineData("1/1/29", 47119)]
    [InlineData("1/1/30", 10959)]
    [InlineData("1-Jan-2020", 43831)]
    [InlineData("1 Jan 2020", 43831)]
    [InlineData("01-January-20", 43831)]
    [InlineData("Jan 1, 2020", 43831)]
    [InlineData("January 1 2020", 43831)]
    [InlineData("jan 1 2020", 43831)]
    [InlineData("Sept 1 2020", 44075)]
    [InlineData("2020-Jan-01", 43831)]
    [InlineData("Jan 2020", 43831)]
    [InlineData("Jan-2020", 43831)]
    [InlineData("1/2020", 43831)]
    [InlineData("29-Feb-1900", 60)]
    [InlineData("1900-02-29", 60)]
    [InlineData("1900-03-01", 61)]
    [InlineData("9999-12-31", 2958465)]
    [InlineData("2024-02-29", 45351)]
    public void Dates(string text, double expected) => Assert.Equal(expected, Parse(text));

    [Theory]
    [InlineData("2026-07-16 4:35 PM", 46219 + 16.5833333333333 / 24)]
    [InlineData("7/7/2026 4:35 PM", 46210 + 16500.0 / 86400 + 0.5)]
    [InlineData("2026-01-01 24:00:00", 46024)]
    [InlineData("2026-01-01 25:00:00", 46024 + 1.0 / 24)]
    [InlineData("Jan 1, 2020 6:00", 43831.25)]
    public void Date_and_time(string text, double expected) => Assert.Equal(expected, Parse(text)!.Value, 9);

    [Theory]
    [InlineData("")]
    [InlineData("abc")]
    [InlineData("1 2")]
    [InlineData("1,23")]
    [InlineData("1.2.3")]
    [InlineData("2026-13-01")]
    [InlineData("2026-00-01")]
    [InlineData("2026-01-00")]
    [InlineData("2026-01-32")]
    [InlineData("2026-04-31")]
    [InlineData("2/30/2020")]
    [InlineData("2025-02-29")]
    [InlineData("2100-02-29")]
    [InlineData("10000-01-01")]
    [InlineData("1789-05-05")]
    [InlineData("9999-13-01")]
    [InlineData("\"2026-01-12\"")]
    [InlineData("2026.01.01")]
    [InlineData("2026-01-01x")]
    [InlineData("-2026-01-01")]
    [InlineData("2026-01-01T12:00")]
    [InlineData("-4:35 AM")]
    [InlineData("4.35")]
    [InlineData("13:00 PM")]
    [InlineData("0:0:86400")]
    [InlineData("4:35 A.M.")]
    [InlineData("4:35 PMx")]
    [InlineData(":35")]
    [InlineData("4::35")]
    [InlineData("1600")]
    [InlineData("10:60")]
    [InlineData("4")]
    [InlineData("Jan")]
    [InlineData("26-01-01")]
    [InlineData("9999-12-31 24:00")]
    public void Not_dates(string text) => Assert.Null(Parse(text));

    [Fact]
    public void Text_without_a_year_needs_the_current_year()
    {
        Assert.Null(Parse("1-Jan"));
        Assert.Null(Parse("3/15"));
        Assert.Null(Parse("Jan 15"));
        Assert.Equal(43831, Parse("1-Jan", year: 2020));
        Assert.Equal(43905, Parse("3/15", year: 2020));
        Assert.Equal(43845, Parse("Jan 15", year: 2020));
        // Not a day of February, so a month and a year, as "1/45" is January 1945.
        Assert.Equal(10990, Parse("2/30", year: 2020));
    }

    [Theory]
    [InlineData("1/2/2020", 43862)]
    [InlineData("26-01-01", 36917)]
    [InlineData("2023-1-2", 44928)]
    [InlineData("1 July 2026", 46204)]
    [InlineData("Jan 2020", 43831)]
    public void Day_first_culture(string text, double expected) => Assert.Equal(expected, Parse(text, EnGb));

    [Theory]
    [InlineData("July 1, 2026")]
    [InlineData("13/13/2020")]
    public void Day_first_culture_rejects(string text) => Assert.Null(Parse(text, EnGb));

    [Theory]
    [InlineData("01.02.2020", 43862)]
    [InlineData("1 января 2020", 43831)]
    [InlineData("1 янв 2020", 43831)]
    [InlineData("1-Jan-2020", 43831)]
    public void Culture_separators_and_month_names(string text, double expected) => Assert.Equal(expected, Parse(text, RuRu));

    [Fact]
    public void Date_system_1904()
    {
        Assert.Equal(42369, Parse("2020-01-01", system: DateSystem.Date1904));
        Assert.Equal(0, Parse("1904-01-01", system: DateSystem.Date1904));
        Assert.Null(Parse("1903-12-31", system: DateSystem.Date1904));
        Assert.Null(Parse("1900-02-29", system: DateSystem.Date1904));
        Assert.Equal(0.5, Parse("12:00", system: DateSystem.Date1904));
    }

    [Theory]
    [InlineData("12:00", 0.5)]
    [InlineData("2020-01-01", 43831)]
    [InlineData("$5", 5)]
    public void Coercion_reads_dates_and_times(string text, double expected)
    {
        Assert.Equal(CellValue.Number(expected), Coercion.ToNumber(text, Invariant, DateSystem.Date1900));
    }

    [Theory]
    [InlineData(" 2024-01-10 ")]
    [InlineData(" 2024-01-10")]
    [InlineData("2024-01-10 ")]
    public void Spaces_around_a_date_are_not_read(string text)
    {
        // As the corpus shows: FLOOR(x, " 2024-01-10 ") is #VALUE! while FLOOR(x, " 10 ") reads 10
        // and DATEVALUE(" 2024-01-10 ") trims (DateTimeFunctions does that before parsing).
        Assert.False(DateText.TryParse(text, System.Globalization.CultureInfo.InvariantCulture, DateSystem.Date1900, out _));
    }
}
