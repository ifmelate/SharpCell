using System;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

// Spans between dates and dates a number of months away: DAYS, DAYS360, DATEDIF, YEARFRAC, EDATE, EOMONTH.
internal static partial class DateTimeFunctions
{
    // EDATE and EOMONTH: more months than this cannot stay inside 1900-9999.
    private const double MaxMonthOffset = 120000;

    private static void RegisterSpans(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("DAYS", 2, 2, Values, Days));
        registry.Add(new FunctionInfo("DAYS360", 2, 3, Values, Days360));
        registry.Add(new FunctionInfo("DATEDIF", 3, 3, Values, DateDif));
        registry.Add(new FunctionInfo("YEARFRAC", 2, 3, Values, YearFrac));
        registry.Add(new FunctionInfo("EDATE", 2, 2, Values, call => MonthsAway(call, endOfMonth: false)));
        registry.Add(new FunctionInfo("EOMONTH", 2, 2, Values, call => MonthsAway(call, endOfMonth: true)));
    }

    // DAYS(end, start): whole days, times of day dropped.
    private static Operand Days(FunctionCall call)
    {
        var end = Serial(call, 0, logicals: true);
        if (end.IsError)
            return end;
        var start = Serial(call, 1, logicals: true);
        if (start.IsError)
            return start;
        return CellValue.Number(Math.Floor(end.AsNumber()) - Math.Floor(start.AsNumber()));
    }

    // Twelve 30-day months. US method (default): a start on the 31st or on the last day of February
    // counts as the 30th; an end on the 31st counts as the 30th when the start is the 30th or later
    // (Excel does not move an end on the last day of February, unlike the NASD rule). European
    // method: every 31st is the 30th. Unlike DAYS, the dates are rounded to the second first.
    private static Operand Days360(FunctionCall call)
    {
        var system = call.Context.DateSystem;
        Span<int> years = stackalloc int[2];
        Span<int> months = stackalloc int[2];
        Span<int> days = stackalloc int[2];
        for (var i = 0; i < 2; i++)
        {
            var serial = Serial(call, i, logicals: true);
            if (serial.IsError)
                return serial;
            var rounded = Math.Round(serial.AsNumber() * 86400, MidpointRounding.AwayFromZero) / 86400;
            if (!DateSerial.TryToDate(rounded, system, out years[i], out months[i], out days[i]))
                return CellValue.Error(ErrorKind.Num);
        }

        var european = call.Boolean(2, false);
        if (european.IsError)
            return european;

        int startDay = days[0], endDay = days[1];
        if (european.AsBoolean())
        {
            startDay = Math.Min(startDay, 30);
            endDay = Math.Min(endDay, 30);
        }
        else
        {
            if (startDay == 31 || (months[0] == 2 && startDay == DateSerial.DaysInMonth(years[0], 2, system)))
                startDay = 30;
            if (endDay == 31 && startDay >= 30)
                endDay = 30;
        }

        return CellValue.Number((years[1] - years[0]) * 360 + (months[1] - months[0]) * 30 + (endDay - startDay));
    }

    // DATEDIF(start, end, unit): complete years ("Y"), months ("M") or days ("D"); days ignoring
    // months and years ("MD"), months ignoring years ("YM"), days ignoring years ("YD").
    private static Operand DateDif(FunctionCall call)
    {
        var system = call.Context.DateSystem;
        var startValue = Serial(call, 0, logicals: true);
        if (startValue.IsError)
            return startValue;
        var endValue = Serial(call, 1, logicals: true);
        if (endValue.IsError)
            return endValue;
        var unitValue = call.Text(2);
        if (unitValue.IsError)
            return unitValue;

        var start = Math.Floor(startValue.AsNumber());
        var end = Math.Floor(endValue.AsNumber());
        if (start > end)
            return CellValue.Error(ErrorKind.Num);
        DateSerial.TryToDate(start, system, out var y1, out var m1, out var d1);
        DateSerial.TryToDate(end, system, out var y2, out var m2, out var d2);

        var wholeMonths = (y2 - y1) * 12 + (m2 - m1) - (d2 < d1 ? 1 : 0);
        switch (unitValue.AsText().ToUpperInvariant())
        {
            case "Y":
                return CellValue.Number(wholeMonths / 12);
            case "M":
                return CellValue.Number(wholeMonths);
            case "D":
                return CellValue.Number(end - start);
            case "YM":
                return CellValue.Number(wholeMonths % 12);
            case "MD":
                // Excel's rule: the start day in the month before the end month, carried over when
                // that month is shorter, which can make the result negative (Jan 31 to Mar 1 is -2).
                if (d1 <= d2)
                    return CellValue.Number(d2 - d1);
                return DateSerial.TryFromDate(y2, m2 - 1, d1, system, out var shifted)
                    ? CellValue.Number(end - shifted)
                    : CellValue.Error(ErrorKind.Num);
            case "YD":
                // The last anniversary of the start on or before the end; February 29 becomes the 28th.
                var anniversary = Anniversary(y2, m1, d1, system);
                if (anniversary > end)
                    anniversary = Anniversary(y2 - 1, m1, d1, system);
                return CellValue.Number(end - anniversary);
            default:
                return CellValue.Error(ErrorKind.Num);
        }
    }

    private static double Anniversary(int year, int month, int day, DateSystem system)
    {
        DateSerial.TryFromDate(year, month, Math.Min(day, DateSerial.DaysInMonth(year, month, system)), system, out var serial);
        return serial;
    }

    // The fraction of a year between two dates by a day-count basis: 0 US 30/360, 1 actual/actual,
    // 2 actual/360, 3 actual/365, 4 European 30/360. The order of the dates does not matter.
    private static Operand YearFrac(FunctionCall call)
    {
        var system = call.Context.DateSystem;
        var startValue = Serial(call, 0, logicals: false);
        if (startValue.IsError)
            return startValue;
        var endValue = Serial(call, 1, logicals: false);
        if (endValue.IsError)
            return endValue;
        var basisValue = call.Has(2) ? Number(call, 2, logicals: false) : CellValue.Number(0);
        if (basisValue.IsError)
            return basisValue;
        var basis = Math.Truncate(basisValue.AsNumber());
        if (basis is < 0 or > 4)
            return CellValue.Error(ErrorKind.Num);

        var start = Math.Floor(startValue.AsNumber());
        var end = Math.Floor(endValue.AsNumber());
        if (start > end)
            (start, end) = (end, start);
        DateSerial.TryToDate(start, system, out var y1, out var m1, out var d1);
        DateSerial.TryToDate(end, system, out var y2, out var m2, out var d2);
        var days = end - start;

        switch (basis)
        {
            case 0:
                var lastOfFebruary1 = m1 == 2 && d1 == DateSerial.DaysInMonth(y1, 2, system);
                var lastOfFebruary2 = m2 == 2 && d2 == DateSerial.DaysInMonth(y2, 2, system);
                if (d1 == 31 && d2 == 31)
                    (d1, d2) = (30, 30);
                else if (d1 == 31)
                    d1 = 30;
                else if (d1 == 30 && d2 == 31)
                    d2 = 30;
                else if (lastOfFebruary1 && lastOfFebruary2)
                    (d1, d2) = (30, 30);
                else if (lastOfFebruary1)
                    d1 = 30;
                return CellValue.Number(Days30(y1, m1, d1, y2, m2, d2) / 360);
            case 1:
                return CellValue.Number(days / ActualYearLength(y1, m1, d1, y2, m2, d2));
            case 2:
                return CellValue.Number(days / 360);
            case 3:
                return CellValue.Number(days / 365);
            default:
                return CellValue.Number(Days30(y1, m1, Math.Min(d1, 30), y2, m2, Math.Min(d2, 30)) / 360);
        }
    }

    private static double Days30(int y1, int m1, int d1, int y2, int m2, int d2) =>
        (y2 - y1) * 360.0 + (m2 - m1) * 30 + (d2 - d1);

    // Actual/actual: the average year length over the years touched when the span is longer than a
    // year; otherwise 366 when a February 29 counts (Excel's rule, which counts one whenever both
    // dates are in the same leap year), else 365.
    private static double ActualYearLength(int y1, int m1, int d1, int y2, int m2, int d2)
    {
        var moreThanAYear = y1 != y2 && (y1 + 1 != y2 || m1 < m2 || (m1 == m2 && d1 < d2));
        if (moreThanAYear)
        {
            var total = 0.0;
            for (var year = y1; year <= y2; year++)
                total += DateTime.IsLeapYear(year) ? 366 : 365;
            return total / (y2 - y1 + 1);
        }

        bool countsFebruary29;
        if (DateTime.IsLeapYear(y1))
            countsFebruary29 = y1 == y2 || m1 <= 2;
        else if (DateTime.IsLeapYear(y2))
            countsFebruary29 = m2 > 2 || (m2 == 2 && d2 == 29);
        else
            countsFebruary29 = false;
        return countsFebruary29 ? 366 : 365;
    }

    // EDATE: the same day a number of months away, or the month's last day when it is shorter.
    // EOMONTH: the last day of the month a number of months away. Months are truncated.
    private static Operand MonthsAway(FunctionCall call, bool endOfMonth)
    {
        var system = call.Context.DateSystem;
        var start = Serial(call, 0, logicals: false);
        if (start.IsError)
            return start;
        var offset = Number(call, 1, logicals: false);
        if (offset.IsError)
            return offset;
        var months = Math.Truncate(offset.AsNumber());
        if (Math.Abs(months) > MaxMonthOffset)
            return CellValue.Error(ErrorKind.Num);

        DateSerial.TryToDate(start.AsNumber(), system, out var year, out var month, out var day);
        var index = year * 12L + (month - 1) + (long)months;
        var targetYear = (int)Math.Floor(index / 12.0);
        var targetMonth = (int)(index - targetYear * 12L) + 1;
        if (targetYear is < 1900 or > 9999)
            return CellValue.Error(ErrorKind.Num);

        var length = DateSerial.DaysInMonth(targetYear, targetMonth, system);
        var targetDay = endOfMonth ? length : Math.Min(day, length);
        return DateSerial.TryFromDate(targetYear, targetMonth, targetDay, system, out var serial)
            ? CellValue.Number(serial)
            : CellValue.Error(ErrorKind.Num);
    }
}
