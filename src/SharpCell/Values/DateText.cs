using System.Globalization;

namespace SharpCell;

/// <summary>
/// Dates and times written as text, as Excel reads them when text meets arithmetic or a date
/// function: <c>"12:00"</c>, <c>"2020-01-01"</c>, <c>"1/2/2020 13:30"</c>. The result is a serial number.
/// </summary>
internal static class DateText
{
    public static bool TryParse(string text, CultureInfo culture, DateSystem dateSystem, out double serial)
    {
        serial = 0;
        return false;
    }
}
