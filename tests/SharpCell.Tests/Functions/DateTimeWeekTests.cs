namespace SharpCell.Tests.Functions;

public class DateTimeWeekTests : DateTimeTestBase
{
    public DateTimeWeekTests()
    {
        // Holidays: 2008-11-26, 2008-12-04, 2009-01-21 with an empty cell between.
        Sheet["H1"].Value = 39778;
        Sheet["H2"].Value = 39786;
        Sheet["H4"].Value = 39834;
        Sheet["J1"].Value = 45667;
        Sheet["J2"].Value = CellValue.Error(ErrorKind.Div0);
        Sheet["J3"].Value = CellValue.Error(ErrorKind.NA);
        Sheet["K1"].Value = true;
    }

    public static TheoryData<string, object> Weekdays => new()
    {
        // 2008-02-14 is a Thursday.
        { "=WEEKDAY(39492)", 5 },
        { "=WEEKDAY(39492,2)", 4 },
        { "=WEEKDAY(39492,3)", 3 },
        { "=WEEKDAY(39492,11)", 4 },
        { "=WEEKDAY(39492,14)", 1 },
        { "=WEEKDAY(39492,16)", 6 },
        { "=WEEKDAY(39492,17)", 5 },
        { "=WEEKDAY(39492.9)", 5 },
        { "=WEEKDAY(\"2008-02-14\")", 5 },
        { "=WEEKDAY(0)", 7 },
        { "=WEEKDAY(1)", 1 },
        { "=WEEKDAY(60)", 4 },
        { "=WEEKDAY(61)", 5 },
        { "=WEEKDAY(39492,4)", "#NUM!" },
        { "=WEEKDAY(39492,0)", "#NUM!" },
        { "=WEEKDAY(-1)", "#NUM!" },
        { "=WEEKDAY(39492,\"x\")", "#VALUE!" },
    };

    [Theory]
    [MemberData(nameof(Weekdays))]
    public void WEEKDAY(string formula, object expected) => Check(formula, expected);

    public static TheoryData<string, object> Weeks => new()
    {
        // 2012-03-09 is a Friday.
        { "=WEEKNUM(DATE(2012,3,9))", 10 },
        { "=WEEKNUM(DATE(2012,3,9),2)", 11 },
        { "=WEEKNUM(DATE(2012,3,9),21)", 10 },
        { "=WEEKNUM(DATE(2012,3,9),15)", 11 },
        { "=WEEKNUM(DATE(2012,3,9),16)", 10 },
        { "=WEEKNUM(DATE(2012,3,9),17)", 10 },
        { "=WEEKNUM(DATE(2012,1,1))", 1 },
        { "=WEEKNUM(DATE(2012,1,1),2)", 1 },
        { "=WEEKNUM(DATE(2012,1,2),2)", 2 },
        { "=WEEKNUM(DATE(2012,12,31))", 53 },
        { "=WEEKNUM(0)", 0 },
        { "=WEEKNUM(DATE(2012,3,9),3)", "#NUM!" },
        { "=WEEKNUM(-1)", "#NUM!" },
        { "=ISOWEEKNUM(DATE(2012,3,9))", 10 },
        { "=ISOWEEKNUM(DATE(2021,1,1))", 53 },
        { "=ISOWEEKNUM(DATE(2018,12,31))", 1 },
        { "=ISOWEEKNUM(DATE(2020,12,31))", 53 },
        { "=ISOWEEKNUM(DATE(9999,12,31))", 52 },
        { "=ISOWEEKNUM(1)", 52 },
        { "=ISOWEEKNUM(0)", 52 },
        { "=ISOWEEKNUM(-1)", "#NUM!" },
    };

    [Theory]
    [MemberData(nameof(Weeks))]
    public void WEEKNUM_and_ISOWEEKNUM(string formula, object expected) => Check(formula, expected);

    public static TheoryData<string, object> NetworkDays => new()
    {
        // Microsoft's documented examples.
        { "=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1))", 110 },
        { "=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),DATE(2012,11,22))", 109 },
        { "=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),{\"2012/11/22\",\"2012/12/4\",\"2013/1/21\"})", 107 },
        { "=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))", 22 },
        { "=NETWORKDAYS.INTL(DATE(2006,2,28),DATE(2006,1,31))", -21 },
        { "=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{\"2006/1/2\",\"2006/1/16\"})", 22 },
        { "=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),\"0010001\",{\"2006/1/2\",\"2006/1/16\"})", 20 },

        { "=NETWORKDAYS(DATE(2025,12,8),DATE(2025,12,8))", 1 },
        { "=NETWORKDAYS(DATE(2025,12,6),DATE(2025,12,7))", 0 },
        { "=NETWORKDAYS(45992.000001,46000.999999)", 7 },
        { "=NETWORKDAYS(DATE(1900,3,1),DATE(9999,12,31))", 2113147 },
        { "=NETWORKDAYS(DATE(2008,10,1),DATE(2009,1,31),H1:H4)", 85 },
        { "=NETWORKDAYS(DATE(2008,10,1),DATE(2009,1,31),{39778,39778})", 87 },
        { "=NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2026,1,1),11)", 314 },
        { "=NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2026,12,31),\"1111111\")", 0 },
        { "=NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2026,12,31),\"0000000\")", 730 },
        { "=NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2026,1,1),0)", "#NUM!" },
        { "=NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2026,1,1),18)", "#NUM!" },
        { "=NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2026,1,1),\"1\")", "#VALUE!" },
        { "=NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2026,1,1),\"0000012\")", "#VALUE!" },
        { "=NETWORKDAYS(1,10,\"abc\")", "#VALUE!" },
        { "=NETWORKDAYS(1,10,K1)", "#VALUE!" },
        { "=NETWORKDAYS(1,10,-2)", "#NUM!" },
        { "=NETWORKDAYS(1,10,J1:J3)", "#DIV/0!" },
        { "=NETWORKDAYS(TRUE,10)", "#VALUE!" },
        { "=NETWORKDAYS(-10,-5)", "#NUM!" },
        { "=NETWORKDAYS(\"2026-06-01\",\"2026-07-01\")", 23 },
    };

    [Theory]
    [MemberData(nameof(NetworkDays))]
    public void NETWORKDAYS(string formula, object expected) => Check(formula, expected);

    public static TheoryData<string, object> Workdays => new()
    {
        // Microsoft's documented examples.
        { "=WORKDAY(DATE(2008,10,1),151)", 39933 },
        { "=WORKDAY(DATE(2008,10,1),151,H1:H4)", 39938 },
        { "=WORKDAY.INTL(DATE(2012,1,1),90,11)", 41013 },
        { "=WORKDAY.INTL(DATE(2012,1,1),30,17)", 40944 },
        { "=WORKDAY.INTL(DATE(2012,1,1),30,0)", "#NUM!" },

        { "=WORKDAY(DATE(2025,12,12),1)", 46006 },
        { "=WORKDAY(DATE(2025,12,13),0)", 46004 },
        { "=WORKDAY(DATE(2025,1,1),-10)", 45644 },
        { "=WORKDAY(DATE(2025,1,13),-1,DATE(2025,1,10))", 45666 },
        { "=WORKDAY(DATE(2025,1,10),3,{45670,45670})", 45673 },
        { "=WORKDAY(DATE(2025,1,10),-0.9999999)", 45666 },
        { "=WORKDAY(DATE(2025,1,10),0.9999999)", 45667 },
        { "=WORKDAY(DATE(2025,1,1),1000000)", 1445658 },
        { "=WORKDAY(8,2)", 10 },
        { "=WORKDAY(DATE(9999,12,31),1)", "#NUM!" },
        { "=WORKDAY(DATE(1900,3,1),-65)", "#NUM!" },
        { "=WORKDAY(DATE(2025,1,10),1E+15)", "#NUM!" },
        { "=WORKDAY(DATE(2025,1,10),TRUE)", "#VALUE!" },
        { "=WORKDAY(DATE(2025,1,10),3,\"abc\")", "#VALUE!" },
        { "=WORKDAY.INTL(DATE(2025,1,10),3,,DATE(2025,1,13))", 45673 },
        { "=WORKDAY.INTL(DATE(2025,1,10),3,\"0000000\",DATE(2025,1,13))", 45671 },
        { "=WORKDAY.INTL(DATE(2025,1,10),3,\"1111110\",DATE(2025,1,13))", 45683 },
        { "=WORKDAY.INTL(DATE(2025,1,10),10,\"1111111\")", "#VALUE!" },
        { "=WORKDAY(\"2025-01-07\",\"10\")", 45678 },
    };

    [Theory]
    [MemberData(nameof(Workdays))]
    public void WORKDAY(string formula, object expected) => Check(formula, expected);

    [Fact]
    public void Date_system_1904()
    {
        Book.DateSystem = DateSystem.Date1904;
        Check("=WEEKDAY(DATE(2008,2,14))", 5);
        Check("=WEEKDAY(0)", 6);
        Check("=ISOWEEKNUM(0)", 53);
        Check("=WEEKNUM(DATE(2012,3,9),2)", 11);
        Check("=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1))", 110);
        Check("=WORKDAY(DATE(2008,10,1),151)-DATE(2008,10,1)", 39933 - 39722);
    }

    [Fact]
    public void Long_spans_are_counted_without_walking_each_day()
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        for (var i = 0; i < 200; i++)
            Check("=NETWORKDAYS(DATE(1900,3,1),DATE(9999,12,31))+WORKDAY(DATE(2025,1,1),1000000)", 2113147 + 1445658);
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(5));
    }
}
