using System.Globalization;
using SharpCell;
using SharpCell.Functions;

namespace SharpCell.Tests.Functions;

/// <summary>Formula → Excel value checks shared by the date and time test classes.</summary>
public abstract class DateTimeTestBase
{
    protected readonly Workbook Book = new();
    protected readonly Worksheet Sheet;

    protected DateTimeTestBase()
    {
        Sheet = Book.AddSheet("S");
        Book.Clock = new FixedClock(new DateTime(2026, 10, 8, 15, 30, 0));
    }

    /// <summary>Expected: a number (compared to 12 significant digits), an error text such as "#NUM!", or other text.</summary>
    protected void Check(string formula, object expected)
    {
        var actual = Book.Evaluate(formula);
        switch (expected)
        {
            case string text when ErrorKinds.TryParse(text, out var kind):
                Assert.Equal(CellValue.Error(kind), actual);
                break;
            case string text:
                Assert.Equal(CellValue.Text(text), actual);
                break;
            case bool flag:
                Assert.Equal(CellValue.Boolean(flag), actual);
                break;
            default:
                var number = Convert.ToDouble(expected, CultureInfo.InvariantCulture);
                Assert.True(actual.Kind == CellValueKind.Number, $"{formula} gave {actual}, expected {number}");
                Assert.Equal(number, actual.AsNumber(), Math.Max(Math.Abs(number), 1) * 1e-12);
                break;
        }
    }

    protected sealed class FixedClock(DateTime local) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => new(local, TimeSpan.Zero);

        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }
}

public class DateTimeFunctionTests : DateTimeTestBase
{
    public static TheoryData<string, object> Dates => new()
    {
        { "=DATE(2020,1,1)", 43831 },
        { "=DATE(1900,1,1)", 1 },
        { "=DATE(1900,2,29)", 60 },
        { "=DATE(1900,3,1)", 61 },
        { "=DATE(1900,1,0)", 0 },
        { "=DATE(120,1,1)", 43831 },
        { "=DATE(0,8,17)", 230 },
        { "=DATE(1800,1,1)", 657439 },
        { "=DATE(2020,13,1)", 44197 },
        { "=DATE(2021,0,1)", 44166 },
        { "=DATE(2020,2,30)", 43891 },
        { "=DATE(2025,1,-1)", 45656 },
        { "=DATE(2025.9,6.9,15.9)", 45823 },
        { "=DATE(\"2025\",\"6\",\"15\")", 45823 },
        { "=DATE(TRUE,6,15)", 532 },
        { "=DATE(9999,12,31)", 2958465 },
        { "=DATE(9999,12,32)", "#NUM!" },
        { "=DATE(10000,1,1)", "#NUM!" },
        { "=DATE(-1,1,1)", "#NUM!" },
        { "=DATE(0,0,0)", "#NUM!" },
        { "=DATE(1900,0,1)", "#NUM!" },
        { "=DATE(1E+300,1,1)", "#NUM!" },
        { "=DATE(2020,1,1E+10)", "#NUM!" },
        { "=DATE(\"abc\",1,1)", "#VALUE!" },
        { "=DATE(1/0,1,1)", "#DIV/0!" },
        { "=DATE(2020,,1)", 43800 },
        { "=SUM(DATE({2020,2021},1,1))", 43831 + 44197 },
    };

    [Theory]
    [MemberData(nameof(Dates))]
    public void DATE(string formula, object expected) => Check(formula, expected);

    public static TheoryData<string, object> Times => new()
    {
        { "=TIME(12,0,0)", 0.5 },
        { "=TIME(3,2,1)", 10921.0 / 86400 },
        { "=TIME(24,2,3)", 0.0014236111111110006 },
        { "=TIME(23,120,0)", 0.04166666666666674 },
        { "=TIME(23,59,60)", 0 },
        { "=TIME(12,-2,10)", 0.49872685185185184 },
        { "=TIME(3.1,20.7,0.9)", 0.1388888888888889 },
        { "=TIME(32767,32767,32767)", 0.4257754629629744 },
        { "=TIME(32768,0,0)", "#NUM!" },
        { "=TIME(0,0,-1)", "#NUM!" },
        { "=TIME(-1,1,1)", "#NUM!" },
        { "=TIME(TRUE,FALSE,TRUE)", 3601.0 / 86400 },
        { "=TIME(\"a\",0,0)", "#VALUE!" },
    };

    [Theory]
    [MemberData(nameof(Times))]
    public void TIME(string formula, object expected) => Check(formula, expected);

    public static TheoryData<string, object> Parts => new()
    {
        { "=YEAR(43831.75)", 2020 },
        { "=MONTH(43831.75)", 1 },
        { "=DAY(43861)", 31 },
        { "=DAY(60)", 29 },
        { "=MONTH(60)", 2 },
        { "=DAY(61)", 1 },
        { "=DAY(0)", 0 },
        { "=MONTH(0)", 1 },
        { "=YEAR(0)", 1900 },
        { "=YEAR(TRUE)", 1900 },
        { "=YEAR(2958465.5)", 9999 },
        { "=YEAR(2958466)", "#NUM!" },
        { "=DAY(-1)", "#NUM!" },
        { "=DAY(\"2020-03-15\")", 15 },
        { "=MONTH(\"45658\")", 1 },
        { "=DAY(\"qwerty\")", "#VALUE!" },
        { "=DAY(\"2025-02-30\")", "#VALUE!" },
        { "=HOUR(0.75)", 18 },
        { "=HOUR(45823.6046875)", 14 },
        { "=MINUTE(45823.6046875)", 30 },
        { "=SECOND(45823.6046875)", 45 },
        { "=MINUTE(\"12:34\")", 34 },
        { "=HOUR(\"25:00\")", 1 },
        { "=SECOND(59.6/86400)", 0 },
        { "=MINUTE(59.6/86400)", 1 },
        { "=HOUR(0.99999999)", 0 },
        { "=HOUR(-0.5)", "#NUM!" },
        { "=HOUR(TRUE)", 0 },
        { "=SECOND(\"x\")", "#VALUE!" },
    };

    [Theory]
    [MemberData(nameof(Parts))]
    public void Calendar_and_clock_parts(string formula, object expected) => Check(formula, expected);

    public static TheoryData<string, object> TextValues => new()
    {
        { "=DATEVALUE(\"2020-01-01\")", 43831 },
        { "=DATEVALUE(\" 2024-01-10 \")", 45301 },
        { "=DATEVALUE(\"1/2/2020 12:00\")", 43832 },
        { "=DATEVALUE(\"2026-01-01 24:00:00\")", 46024 },
        { "=DATEVALUE(\"29-Feb-1900\")", 60 },
        { "=DATEVALUE(\"Jan 1, 2020\")", 43831 },
        { "=DATEVALUE(\"1-Jan\")", 46023 },
        { "=DATEVALUE(\"5-JUL\")", 46208 },
        { "=DATEVALUE(\"12:00\")", 0 },
        { "=DATEVALUE(\"2025-02-29\")", "#VALUE!" },
        { "=DATEVALUE(\"8\")", "#VALUE!" },
        { "=DATEVALUE(43831)", "#VALUE!" },
        { "=DATEVALUE(TRUE)", "#VALUE!" },
        { "=DATEVALUE(\"\")", "#VALUE!" },
        { "=DATEVALUE(#N/A)", "#N/A" },
        { "=TIMEVALUE(\"6:00 PM\")", 0.75 },
        { "=TIMEVALUE(\"4:35 AM\")", 0.1909722222222222 },
        { "=TIMEVALUE(\"24:00\")", 0 },
        { "=TIMEVALUE(\"25:00\")", 1.0 / 24 },
        { "=TIMEVALUE(\"7/7/2026 4:35 PM\")", 0.6909722222189885 },
        { "=TIMEVALUE(\"2020-01-01\")", 0 },
        { "=TIMEVALUE(\"13:00 PM\")", "#VALUE!" },
        { "=TIMEVALUE(0.5)", "#VALUE!" },
        { "=TIMEVALUE(\"1600\")", "#VALUE!" },
    };

    [Theory]
    [MemberData(nameof(TextValues))]
    public void DATEVALUE_and_TIMEVALUE(string formula, object expected) => Check(formula, expected);

    public static TheoryData<string, object> TextInArithmetic => new()
    {
        { "=--\"12:00\"", 0.5 },
        { "=\"2020-01-01\"+1", 43832 },
        { "=MAX(\"12:00\",0.4)", 0.5 },
        { "=\"25:00\"*24", 25 },
        { "=\"1/2/2020\"-\"1/1/2020\"", 1 },
        { "=\"2/30/2020\"+0", "#VALUE!" },

        // The current year needs the clock, which coercion does not have.
        { "=\"1-Jan\"+0", "#VALUE!" },
    };

    [Theory]
    [MemberData(nameof(TextInArithmetic))]
    public void Text_dates_and_times_are_numbers(string formula, object expected) => Check(formula, expected);

    [Fact]
    public void TODAY_is_the_date_of_NOW_and_volatile()
    {
        var today = (new DateTime(2026, 10, 8) - new DateTime(1899, 12, 30)).TotalDays;
        Check("=TODAY()", today);
        var now = (new DateTime(2026, 10, 8, 15, 30, 0) - new DateTime(1899, 12, 30)).TotalDays;
        Check("=NOW()-TODAY()", now - today);
        Book.DateSystem = DateSystem.Date1904;
        Check("=TODAY()", today - 1462);
        Assert.True(FunctionRegistry.Default.All.Single(f => f.Name == "TODAY").IsVolatile);
    }

    [Fact]
    public void Date_system_1904()
    {
        Book.DateSystem = DateSystem.Date1904;
        Check("=DATE(2020,1,1)", 42369);
        Check("=DATE(1904,1,1)", 0);
        Check("=DATE(1903,12,31)", "#NUM!");
        Check("=YEAR(0)", 1904);
        Check("=DAY(42369)", 1);
        Check("=DATEVALUE(\"2020-01-01\")", 42369);
        Check("=DATEVALUE(\"1900-02-29\")", "#VALUE!");
    }

    [Fact]
    public void Text_dates_follow_the_workbook_culture()
    {
        Book.Culture = CultureInfo.GetCultureInfo("en-GB");
        Check("=DATEVALUE(\"1/2/2020\")", 43862);
        Check("=DATEVALUE(\"26-01-01\")", 36917);
        Check("=DATEVALUE(\"July 1, 2026\")", "#VALUE!");
        Book.Culture = CultureInfo.GetCultureInfo("ru-RU");
        Check("=DATEVALUE(\"01.02.2020\")", 43862);
    }

    [Fact]
    public void Arrays_are_mapped()
    {
        Sheet["A1"].Value = 43831;
        Sheet["A2"].Value = 43862;
        Check("=SUM(MONTH(A1:A2))", 3);
        Check("=SUM(DAY({43831,43832}))", 3);
    }
}
