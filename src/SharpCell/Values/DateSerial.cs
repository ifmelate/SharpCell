using System;

namespace SharpCell;

/// <summary>Excel date serial numbers. DateTime is used only at the boundary, never for date math.</summary>
internal static class DateSerial
{
    private static readonly DateTime Epoch1900 = new(1899, 12, 31);
    private static readonly DateTime Epoch1904 = new(1904, 1, 1);
    private static readonly DateTime FirstRealMarch1900 = new(1900, 3, 1);

    /// <summary>
    /// Serial number with the time of day as the fraction. In the 1900 system serial 60 is the
    /// non-existent 1900-02-29, so dates from 1900-03-01 on are one day later than a plain count.
    /// </summary>
    public static double FromDateTime(DateTime value, DateSystem system)
    {
        if (system == DateSystem.Date1904)
            return (value - Epoch1904).TotalDays;

        var days = (value - Epoch1900).TotalDays;
        return value >= FirstRealMarch1900 ? days + 1 : days;
    }

    /// <summary>The last day Excel knows, 9999-12-31, as a serial number.</summary>
    public static double MaxSerial(DateSystem system) => system == DateSystem.Date1904 ? 2957003 : 2958465;

    /// <summary>
    /// The calendar date of a serial number (its whole part), by Excel's calendar: in the 1900
    /// system serial 0 is 1900-01-00 and serial 60 is the non-existent 1900-02-29.
    /// </summary>
    public static bool TryToDate(double serial, DateSystem system, out int year, out int month, out int day)
    {
        year = month = day = 0;
        if (!(serial >= 0) || serial >= MaxSerial(system) + 1)
            return false;

        var days = (int)Math.Floor(serial);
        DateTime date;
        if (system == DateSystem.Date1904)
        {
            date = Epoch1904.AddDays(days);
        }
        else if (days == 0)
        {
            (year, month, day) = (1900, 1, 0);
            return true;
        }
        else if (days == 60)
        {
            (year, month, day) = (1900, 2, 29);
            return true;
        }
        else
        {
            date = Epoch1900.AddDays(days < 60 ? days : days - 1);
        }

        (year, month, day) = (date.Year, date.Month, date.Day);
        return true;
    }

    /// <summary>
    /// The serial number of a date as the DATE function builds it: years 0 to 1899 count from 1900,
    /// months and days beyond their range carry over (month 13 is January of the next year, day 0
    /// the last day of the previous month), and 1900-02-29 exists. False outside the calendar.
    /// </summary>
    public static bool TryFromDate(long year, long month, long day, DateSystem system, out double serial)
    {
        serial = 0;
        if (year is < 0 or > 9999)
            return false;
        if (year < 1900)
            year += 1900;

        // Months carry over into years, rounding down for negative month numbers.
        var months = year * 12 + (month - 1);
        var normalizedYear = Math.DivRem(months, 12, out var remainder);
        if (remainder < 0)
        {
            remainder += 12;
            normalizedYear--;
        }

        if (normalizedYear is < 1 or > 9999)
            return false;

        var first = new DateTime((int)normalizedYear, (int)remainder + 1, 1);
        double firstSerial = system == DateSystem.Date1904
            ? (first - Epoch1904).TotalDays
            : (first - Epoch1900).TotalDays + (first >= FirstRealMarch1900 ? 1 : 0);
        serial = firstSerial + day - 1;
        return serial >= 0 && serial <= MaxSerial(system);
    }

    /// <summary>Days in a month by Excel's calendar: February 1900 has 29 days in the 1900 system.</summary>
    public static int DaysInMonth(int year, int month, DateSystem system) =>
        year == 1900 && month == 2 && system == DateSystem.Date1900 ? 29 : DateTime.DaysInMonth(year, month);

    /// <summary>
    /// The serial number of a calendar date that must exist as written (no carry-over, year as
    /// written): what a typed date such as "2020-02-30" is checked against. False outside the calendar.
    /// </summary>
    public static bool TryFromCalendar(int year, int month, int day, DateSystem system, out double serial)
    {
        serial = 0;
        if (year is < 1900 or > 9999 || month is < 1 or > 12 || day < 1 || day > DaysInMonth(year, month, system))
            return false;
        return TryFromDate(year, month, day, system, out serial);
    }

    /// <summary>
    /// Day of the week of a serial number's whole part, Monday = 0 to Sunday = 6, by Excel's
    /// calendar: in the 1900 system serial 1 (1900-01-01) is a Sunday, as Excel counts it, although
    /// the real day was a Monday; from 1900-03-01 on the days are right.
    /// </summary>
    public static int MondayBasedWeekday(double serial, DateSystem system)
    {
        var days = (long)Math.Floor(serial) + (system == DateSystem.Date1904 ? 1462 : 0);
        return (int)(((days + 5) % 7 + 7) % 7);
    }

    /// <summary>Hour, minute and second of a serial number's fraction, rounded to the nearest second as Excel does.</summary>
    public static (int Hour, int Minute, int Second) TimeOfDay(double serial)
    {
        var fraction = serial - Math.Floor(serial);
        var seconds = (long)Math.Round(fraction * 86400, MidpointRounding.AwayFromZero);
        if (seconds >= 86400)
            seconds = 0;
        return ((int)(seconds / 3600), (int)(seconds / 60 % 60), (int)(seconds % 60));
    }
}
