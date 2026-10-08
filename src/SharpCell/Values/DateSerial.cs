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
}
