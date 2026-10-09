namespace SharpCell.Tests.Functions;

public class DateTimeSpanTests : DateTimeTestBase
{
    public DateTimeSpanTests()
    {
        Sheet["A1"].Value = true;
    }

    public static TheoryData<string, object> Days => new()
    {
        { "=DAYS(DATE(2021,3,15),DATE(2021,2,1))", 42 },
        { "=DAYS(\"2021-03-15\",\"2021-02-01\")", 42 },
        { "=DAYS(DATE(2025,1,1),DATE(2025,1,10))", -9 },
        { "=DAYS(46000.999999,45992.000001)", 8 },
        { "=DAYS(TRUE,FALSE)", 1 },
        { "=DAYS(10,)", 10 },
        { "=DAYS(-5,-10)", "#NUM!" },
        { "=DAYS(2958466,1)", "#NUM!" },
        { "=DAYS(\"abc\",1)", "#VALUE!" },
        { "=DAYS(1/0,1)", "#DIV/0!" },
    };

    [Theory]
    [MemberData(nameof(Days))]
    public void DAYS(string formula, object expected) => Check(formula, expected);

    public static TheoryData<string, object> Days360 => new()
    {
        { "=DAYS360(DATE(2025,1,1),DATE(2025,1,31))", 30 },
        { "=DAYS360(DATE(2025,1,1),DATE(2025,1,31),TRUE)", 29 },
        { "=DAYS360(DATE(2025,1,1),DATE(2025,1,31),\"false\")", 30 },
        { "=DAYS360(DATE(2025,1,1),DATE(2025,1,31),1)", 29 },
        { "=DAYS360(DATE(2025,1,1),DATE(2025,1,31),\"1\")", "#VALUE!" },
        { "=DAYS360(DATE(2025,1,1),DATE(2026,1,1))", 360 },
        { "=DAYS360(DATE(2025,7,30),DATE(2025,7,31))", 0 },
        { "=DAYS360(DATE(2025,1,31),DATE(2025,2,28))", 28 },
        { "=DAYS360(DATE(2025,2,28),DATE(2025,3,31))", 30 },
        { "=DAYS360(DATE(2025,2,28),DATE(2025,2,28))", -2 },
        { "=DAYS360(DATE(2024,2,29),DATE(2025,2,28))", 358 },
        { "=DAYS360(DATE(2024,2,28),DATE(2024,3,1))", 3 },
        { "=DAYS360(DATE(2025,1,1),DATE(2025,12,31))", 360 },
        { "=DAYS360(DATE(2025,1,10),)", -45010 },
        { "=DAYS360(DATE(2025,1,1)+0.9999999,DATE(2025,1,1))", -1 },
        { "=DAYS360(DATE(2025,1,1)+0.9999,DATE(2025,1,1))", 0 },
        { "=DAYS360(45992.000001,46000.999999)", 9 },
        { "=DAYS360(DATE(1990,5,12),DATE(9990,5,12))", 2880000 },
        { "=DAYS360(-10,-5)", "#NUM!" },
    };

    [Theory]
    [MemberData(nameof(Days360))]
    public void DAYS360(string formula, object expected) => Check(formula, expected);

    public static TheoryData<string, object> DateDif => new()
    {
        { "=DATEDIF(DATE(2020,1,15),DATE(2021,3,10),\"Y\")", 1 },
        { "=DATEDIF(DATE(2020,1,15),DATE(2021,3,10),\"M\")", 13 },
        { "=DATEDIF(DATE(2020,1,15),DATE(2021,3,10),\"D\")", 420 },
        { "=DATEDIF(DATE(2020,1,15),DATE(2021,3,10),\"YM\")", 1 },
        { "=DATEDIF(DATE(2020,1,15),DATE(2021,3,10),\"MD\")", 23 },
        { "=DATEDIF(DATE(2020,1,15),DATE(2021,3,10),\"YD\")", 54 },
        { "=DATEDIF(DATE(2020,1,15),DATE(2021,3,10),\"yd\")", 54 },
        { "=DATEDIF(DATE(2024,1,1),DATE(2024,12,31),\"MD\")", 30 },
        { "=DATEDIF(DATE(2024,1,1),DATE(2024,12,31),\"YD\")", 365 },
        { "=DATEDIF(DATE(2015,1,31),DATE(2015,3,1),\"MD\")", -2 },
        { "=DATEDIF(DATE(2021,1,31),DATE(2021,2,28),\"MD\")", 28 },
        { "=DATEDIF(DATE(2020,2,29),DATE(2021,2,28),\"YD\")", 0 },
        { "=DATEDIF(DATE(2020,2,29),DATE(2021,3,1),\"YD\")", 1 },
        { "=DATEDIF(DATE(2020,2,29),DATE(2021,2,28),\"Y\")", 0 },
        { "=DATEDIF(DATE(2025,1,10)+0.8,DATE(2025,1,10)+0.6,\"D\")", 0 },
        { "=DATEDIF(DATE(2025,1,10),DATE(2025,1,1),\"D\")", "#NUM!" },
        { "=DATEDIF(1,2,\"X\")", "#NUM!" },
        { "=DATEDIF(-1,2,\"D\")", "#NUM!" },
        { "=DATEDIF(\"2024-01-10\",\"2026-01-10\",\"D\")", 731 },
        { "=DATEDIF(1,2,#N/A)", "#N/A" },
    };

    [Theory]
    [MemberData(nameof(DateDif))]
    public void DATEDIF(string formula, object expected) => Check(formula, expected);

    public static TheoryData<string, object> YearFrac => new()
    {
        // Microsoft's documented examples.
        { "=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))", 0.58055555555555556 },
        { "=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)", 0.57650273224043716 },
        { "=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)", 0.57808219178082192 },

        { "=YEARFRAC(DATE(2025,1,10),DATE(2025,1,1))", 0.025 },
        { "=YEARFRAC(DATE(2024,2,29),DATE(2025,2,28))", 1 },
        { "=YEARFRAC(DATE(2025,1,31),DATE(2025,2,28))", 28.0 / 360 },
        { "=YEARFRAC(DATE(2024,1,1),DATE(2024,12,31),1)", 365.0 / 366 },
        { "=YEARFRAC(DATE(2024,1,1),DATE(2024,12,31),2)", 365.0 / 360 },
        { "=YEARFRAC(DATE(2024,1,1),DATE(2024,12,31),4)", 359.0 / 360 },
        { "=YEARFRAC(DATE(2008,3,1),DATE(2008,8,31),3)", 0.5013698630136987 },
        { "=YEARFRAC(DATE(1998,8,8),DATE(2024,4,4),1)", 25.65574934090448 },
        { "=YEARFRAC(DATE(2023,3,3),DATE(2024,2,28),1)", 0.9917808219178083 },
        { "=YEARFRAC(DATE(2022,2,2),DATE(2024,2,29),1)", 2.0720802919708032 },
        { "=YEARFRAC(DATE(1995,5,5),DATE(2026,6,6),1)", 31.08829568788501 },
        { "=YEARFRAC(DATE(2025,1,1),DATE(2025,1,10),1.9)", 9.0 / 365 },
        { "=YEARFRAC(1,2,5)", "#NUM!" },
        { "=YEARFRAC(1,2,-1)", "#NUM!" },
        { "=YEARFRAC(1,2,A1)", "#VALUE!" },
        { "=YEARFRAC(A1,2)", "#VALUE!" },
        { "=YEARFRAC(-10,-5)", "#NUM!" },
        { "=YEARFRAC(\"2024-01-10\",\"2026-01-10\")", 2 },
    };

    [Theory]
    [MemberData(nameof(YearFrac))]
    public void YEARFRAC(string formula, object expected) => Check(formula, expected);

    public static TheoryData<string, object> MonthsAway => new()
    {
        { "=EDATE(DATE(2011,1,15),1)", 40589 },
        { "=EDATE(DATE(2011,1,15),-1)", 40527 },
        { "=EDATE(DATE(2024,1,31),1)", 45351 },
        { "=EDATE(DATE(2024,3,31),-1)", 45351 },
        { "=EDATE(DATE(2024,2,29),12)", 45716 },
        { "=EDATE(32112.5,82.9)", 34608 },
        { "=EDATE(DATE(1987,12,1),-235.5)", 24959 },
        { "=EDATE(DATE(1900,3,1),-7)", "#NUM!" },
        { "=EDATE(2958465,1)", "#NUM!" },
        { "=EDATE(DATE(1987,11,30),100000)", "#NUM!" },
        { "=EDATE(DATE(1987,11,30),1E+300)", "#NUM!" },
        { "=EDATE(TRUE,1)", "#VALUE!" },
        { "=EDATE(DATE(1987,11,30),A1)", "#VALUE!" },
        { "=EDATE(\"2025-06-15\",\"2\")", 45884 },
        { "=EDATE(\"x\",1)", "#VALUE!" },
        { "=EOMONTH(DATE(2011,1,1),1)", 40602 },
        { "=EOMONTH(DATE(2011,1,1),-3)", 40482 },
        { "=EOMONTH(DATE(2023,3,2),-4.7)", 44895 },
        { "=EOMONTH(DATE(2024,1,31),1)", 45351 },
        { "=EOMONTH(DATE(9999,12,31),0)", 2958465 },
        { "=EOMONTH(DATE(9999,12,31),1)", "#NUM!" },
        { "=EOMONTH(-1,0)", "#NUM!" },
        { "=EOMONTH(DATE(2023,3,2),)", 45016 },
    };

    [Theory]
    [MemberData(nameof(MonthsAway))]
    public void EDATE_and_EOMONTH(string formula, object expected) => Check(formula, expected);

    [Fact]
    public void Date_system_1904()
    {
        Book.DateSystem = DateSystem.Date1904;
        Check("=EDATE(DATE(2024,1,31),1)", 45351 - 1462);
        Check("=EOMONTH(DATE(1904,1,15),-1)", "#NUM!");
        Check("=DAYS360(DATE(2024,2,29),DATE(2025,2,28))", 358);
        Check("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)", 0.57650273224043716);
        Check("=DATEDIF(DATE(2020,1,15),DATE(2021,3,10),\"YD\")", 54);
    }
}
