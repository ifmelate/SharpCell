using System;
using System.Data;
using System.Globalization;

namespace SharpCell;

/// <summary>Turns .NET objects, such as the properties of a record being loaded into a sheet, into cell values.</summary>
internal static class ObjectValues
{
    /// <summary>
    /// The cell value for an object, and the number format its cell takes, or null to keep the cell's format.
    /// Throws <see cref="ArgumentException"/> for a number Excel cannot store (NaN, infinity) and lets
    /// <see cref="CellValue.DateTime"/>'s <see cref="ArgumentOutOfRangeException"/> through for a date before the date system starts.
    /// </summary>
    public static CellValue Convert(object? value, Workbook workbook, out string? format)
    {
        format = null;
        var culture = workbook.Culture;
        switch (value)
        {
            case null or DBNull:
                return CellValue.Empty;
            case CellValue cell:
                return cell;
            case string text:
                return CellValue.Text(text);
            case char c:
                return CellValue.Text(c.ToString());
            case bool flag:
                return CellValue.Boolean(flag);
            case byte or sbyte or short or ushort or int or uint or long or ulong:
                return CellValue.Number(System.Convert.ToDouble(value, CultureInfo.InvariantCulture));
            case decimal number:
                return CellValue.Number((double)number);
            case double number:
                return Finite(number, value);
            case float number:
                return Finite(number, value);
            case DateTime date:
                return Date(date, workbook, out format);
            case DateOnly date:
                return Date(date.ToDateTime(TimeOnly.MinValue), workbook, out format);
            case DateTimeOffset date:
                // Excel has no offsets: the clock time the value states.
                return Date(date.DateTime, workbook, out format);
            case TimeOnly time:
                format = CellInput.TimeCode(culture, seconds: HasSeconds(time.Ticks));
                return CellValue.Number((double)time.Ticks / TimeSpan.TicksPerDay);
            case TimeSpan span:
                format = "[h]:mm:ss";
                return CellValue.Number(span.TotalDays);
            case Enum:
                return CellValue.Text(value.ToString()!);
            case IFormattable formattable:
                return CellValue.Text(formattable.ToString(null, culture));
            default:
                return CellValue.Text(value.ToString() ?? "");
        }
    }

    private static CellValue Finite(double number, object value) =>
        double.IsFinite(number)
            ? CellValue.Number(number)
            : throw new ArgumentException($"A {value.GetType().Name} of {number.ToString(CultureInfo.InvariantCulture)} cannot be stored in a cell.");

    private static CellValue Date(DateTime date, Workbook workbook, out string format)
    {
        var result = CellValue.DateTime(date, workbook.DateSystem);
        var culture = workbook.Culture;
        var dateCode = CellInput.DateCode(culture);
        // The time as a typed date-time shows it: hours and minutes, seconds only when there are some.
        format = date.TimeOfDay == TimeSpan.Zero
            ? dateCode
            : dateCode + " h:mm" + (HasSeconds(date.TimeOfDay.Ticks) ? ":ss" : "");
        return result;
    }

    private static bool HasSeconds(long ticks) => ticks % TimeSpan.TicksPerMinute != 0;
}
