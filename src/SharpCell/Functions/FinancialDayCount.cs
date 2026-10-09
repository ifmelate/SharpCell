using System;

namespace SharpCell.Functions;

/// <summary>
/// A calendar date for the day-count arithmetic of the securities functions. It is plain
/// year/month/day with no range check, so stepping a coupon schedule never fails; the day of
/// the phantom 1900-02-29 (serial 60) is kept as it is.
/// </summary>
internal readonly record struct FinancialDate(int Year, int Month, int Day) : IComparable<FinancialDate>
{
    public int CompareTo(FinancialDate other) =>
        Year != other.Year ? Year.CompareTo(other.Year)
        : Month != other.Month ? Month.CompareTo(other.Month)
        : Day.CompareTo(other.Day);

    public static bool operator <(FinancialDate left, FinancialDate right) => left.CompareTo(right) < 0;

    public static bool operator >(FinancialDate left, FinancialDate right) => left.CompareTo(right) > 0;

    public static bool operator <=(FinancialDate left, FinancialDate right) => left.CompareTo(right) <= 0;

    public static bool operator >=(FinancialDate left, FinancialDate right) => left.CompareTo(right) >= 0;

    public static FinancialDate Min(FinancialDate a, FinancialDate b) => a <= b ? a : b;

    public static FinancialDate Max(FinancialDate a, FinancialDate b) => a >= b ? a : b;

    public bool IsLastDayOfMonth => Day >= FinancialDayCount.DaysInMonth(Year, Month);

    public bool IsLastDayOfFebruary => Month == 2 && IsLastDayOfMonth;

    /// <summary>
    /// The date <paramref name="months"/> later, the day clamped to the target month's length
    /// (Jan 31 + 1 month is Feb 28); with <paramref name="toMonthEnd"/>, the last day of that month.
    /// </summary>
    public FinancialDate AddMonths(int months, bool toMonthEnd = false)
    {
        var total = Year * 12 + (Month - 1) + months;
        var year = (int)Math.Floor(total / 12.0);
        var month = total - year * 12 + 1;
        var length = FinancialDayCount.DaysInMonth(year, month);
        return new FinancialDate(year, month, toMonthEnd ? length : Math.Min(Day, length));
    }

    /// <summary>
    /// The serial number in the 1900 date system, counting the phantom 1900-02-29 as Excel does.
    /// Differences of these numbers are actual days in either date system.
    /// </summary>
    public long Serial1900
    {
        get
        {
            var days = FinancialDayCount.DaysFromCivil(Year, Month, Day) - FinancialDayCount.DaysFromCivil(1899, 12, 31);
            return this >= new FinancialDate(1900, 3, 1) ? days + 1 : days;
        }
    }

    /// <summary>The serial number of the date in <paramref name="system"/>.</summary>
    public double ToSerial(DateSystem system) => system == DateSystem.Date1904
        ? FinancialDayCount.DaysFromCivil(Year, Month, Day) - FinancialDayCount.DaysFromCivil(1904, 1, 1)
        : Serial1900;
}

/// <summary>
/// Excel's day-count bases for securities: 0 US (NASD) 30/360, 1 actual/actual, 2 actual/360,
/// 3 actual/365, 4 European 30/360; and the coupon schedule, which runs back from maturity in
/// steps of 12/frequency months, kept on month ends when maturity is on one.
/// </summary>
internal static class FinancialDayCount
{
    public static bool IsLeapYear(int year) => (year % 4 == 0 && year % 100 != 0) || year % 400 == 0;

    public static int DaysInMonth(int year, int month) => month switch
    {
        2 => IsLeapYear(year) ? 29 : 28,
        4 or 6 or 9 or 11 => 30,
        _ => 31,
    };

    // Days since 0000-03-01 of the proleptic Gregorian calendar; linear in the day, so a day
    // beyond the month's end (the phantom 1900-02-29) counts on into the next month.
    public static long DaysFromCivil(int year, int month, int day)
    {
        long y = month <= 2 ? year - 1 : year;
        var era = (long)Math.Floor(y / 400.0);
        var yearOfEra = y - era * 400;
        var monthIndex = month > 2 ? month - 3 : month + 9;
        var dayOfYear = (153 * monthIndex + 2) / 5 + day - 1;
        var dayOfEra = yearOfEra * 365 + yearOfEra / 4 - yearOfEra / 100 + dayOfYear;
        return era * 146097 + dayOfEra;
    }

    public static double Actual(FinancialDate start, FinancialDate end) => end.Serial1900 - start.Serial1900;

    private static double Diff360(int startDay, int startMonth, int startYear, int endDay, int endMonth, int endYear) =>
        (endYear - startYear) * 360.0 + (endMonth - startMonth) * 30.0 + (endDay - startDay);

    /// <summary>
    /// US (NASD) 30/360 days. Excel's securities functions use it in two variants: the plain one,
    /// which moves the end date only when the start is a month end, and one that moves both.
    /// </summary>
    public static double Days360Us(FinancialDate start, FinancialDate end, bool modifyBoth = false)
    {
        var startDay = start.Day;
        var endDay = end.Day;
        if (end.IsLastDayOfFebruary && (start.IsLastDayOfFebruary || modifyBoth))
            endDay = 30;
        if (endDay == 31 && (startDay >= 30 || modifyBoth))
            endDay = 30;
        if (startDay == 31)
            startDay = 30;
        if (start.IsLastDayOfFebruary)
            startDay = 30;
        return Diff360(startDay, start.Month, start.Year, endDay, end.Month, end.Year);
    }

    /// <summary>European 30/360 days: the 31st counts as the 30th.</summary>
    public static double Days360Eu(FinancialDate start, FinancialDate end) =>
        Diff360(Math.Min(start.Day, 30), start.Month, start.Year, Math.Min(end.Day, 30), end.Month, end.Year);

    // Actual/365 in a denominator: whole years count 365 days and February ends on the 28th.
    private static double Days365(FinancialDate start, FinancialDate end)
    {
        var startDay = start.Month == 2 && start.Day > 28 ? 28 : start.Day;
        var endDay = end.Month == 2 && end.Day > 28 ? 28 : end.Day;
        var shifted = new FinancialDate(end.Year, start.Month, startDay);
        return (end.Year - start.Year) * 365.0 + Actual(shifted, new FinancialDate(end.Year, end.Month, endDay));
    }

    /// <summary>
    /// Days from <paramref name="start"/> to <paramref name="end"/> by the basis. The actual/360 and
    /// actual/365 bases count actual days in a numerator but 30/360 and 365-day years in a denominator.
    /// </summary>
    public static double DaysBetween(FinancialDate start, FinancialDate end, int basis, bool numerator = true) => basis switch
    {
        0 => Days360Us(start, end),
        2 => numerator ? Actual(start, end) : Days360Us(start, end),
        3 => numerator ? Actual(start, end) : Days365(start, end),
        4 => Days360Eu(start, end),
        _ => Actual(start, end),
    };

    /// <summary>The year length the basis divides by; for actual/actual it depends on the dates.</summary>
    public static double DaysInYear(FinancialDate start, FinancialDate end, int basis)
    {
        switch (basis)
        {
            case 1:
                if (!WithinAYear(start, end))
                {
                    var years = end.Year - start.Year + 1;
                    return Actual(new FinancialDate(start.Year, 1, 1), new FinancialDate(end.Year + 1, 1, 1)) / years;
                }

                return (start.Year == end.Year && IsLeapYear(start.Year))
                       || (end.Month == 2 && end.Day == 29)
                       || Feb29Between(start, end) ? 366 : 365;
            case 3:
                return 365;
            default:
                return 360;
        }
    }

    /// <summary>The fraction of a year from <paramref name="start"/> to a later <paramref name="end"/>, as YEARFRAC.</summary>
    public static double YearFraction(FinancialDate start, FinancialDate end, int basis) =>
        DaysBetween(start, end, basis) / DaysInYear(start, end, basis);

    // The same year, or the next one up to the same month and day.
    private static bool WithinAYear(FinancialDate start, FinancialDate end) =>
        start.Year == end.Year
        || (end.Year == start.Year + 1 && (start.Month > end.Month || (start.Month == end.Month && start.Day >= end.Day)));

    private static bool Feb29Between(FinancialDate start, FinancialDate end)
    {
        var march1 = new FinancialDate(start.Year, 3, 1);
        if (IsLeapYear(start.Year) && start < march1 && end >= march1)
            return true;
        march1 = new FinancialDate(end.Year, 3, 1);
        return IsLeapYear(end.Year) && end >= march1 && start < march1;
    }

    // Coupon schedule.

    public static int MonthsPerPeriod(int frequency) => 12 / frequency;

    /// <summary>
    /// The coupon dates around settlement: the last one on or before it and the first one after it,
    /// stepping back from maturity.
    /// </summary>
    public static (FinancialDate Previous, FinancialDate Next) CouponDates(FinancialDate settlement, FinancialDate maturity, int frequency)
    {
        var months = MonthsPerPeriod(frequency);
        var monthEnd = maturity.IsLastDayOfMonth;
        var previous = maturity;
        var next = maturity;
        while (previous > settlement)
        {
            next = previous;
            previous = previous.AddMonths(-months, monthEnd);
        }

        return (previous, next);
    }

    /// <summary>The number of coupons payable after settlement up to and including maturity.</summary>
    public static int CouponCount(FinancialDate settlement, FinancialDate maturity, int frequency)
    {
        var previous = CouponDates(settlement, maturity, frequency).Previous;
        var months = (maturity.Year - previous.Year) * 12 + maturity.Month - previous.Month;
        return months * frequency / 12;
    }

    /// <summary>COUPDAYS: the length of the coupon period that contains settlement.</summary>
    public static double CouponDays(FinancialDate settlement, FinancialDate maturity, int frequency, int basis)
    {
        switch (basis)
        {
            case 1:
                var (previous, next) = CouponDates(settlement, maturity, frequency);
                return Actual(previous, next);
            case 3:
                return 365.0 / frequency;
            default:
                return 360.0 / frequency;
        }
    }

    /// <summary>COUPDAYBS: days from the start of the coupon period to settlement.</summary>
    public static double CouponDaysBeforeSettlement(FinancialDate settlement, FinancialDate maturity, int frequency, int basis)
    {
        var previous = CouponDates(settlement, maturity, frequency).Previous;
        return basis switch
        {
            0 => Days360Us(previous, settlement),
            4 => Days360Eu(previous, settlement),
            _ => Actual(previous, settlement),
        };
    }

    /// <summary>COUPDAYSNC: days from settlement to the next coupon date.</summary>
    public static double CouponDaysAfterSettlement(FinancialDate settlement, FinancialDate maturity, int frequency, int basis)
    {
        var (previous, next) = CouponDates(settlement, maturity, frequency);
        return basis switch
        {
            0 => Days360Us(previous, next, modifyBoth: true) - Days360Us(previous, settlement),
            4 => Days360Eu(settlement, next),
            _ => Actual(settlement, next),
        };
    }
}
