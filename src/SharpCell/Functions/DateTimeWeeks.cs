using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

// Weeks and working days: WEEKDAY, WEEKNUM, ISOWEEKNUM, NETWORKDAYS, WORKDAY and their .INTL forms.
// Weekdays are Monday-based (Monday = 0) throughout, and a weekend is a 7-bit mask, bit 0 Monday.
internal static partial class DateTimeFunctions
{
    // Saturday and Sunday.
    private const int DefaultWeekend = 0b110_0000;

    private const int WholeWeek = 0b111_1111;

    private static void RegisterWeeks(FunctionRegistry registry)
    {
        ArgumentKind[] withHolidays = [ArgumentKind.Value, ArgumentKind.Value, ArgumentKind.Any];
        ArgumentKind[] withWeekendAndHolidays = [ArgumentKind.Value, ArgumentKind.Value, ArgumentKind.Value, ArgumentKind.Any];
        registry.Add(new FunctionInfo("WEEKDAY", 1, 2, Values, Weekday));
        registry.Add(new FunctionInfo("WEEKNUM", 1, 2, Values, WeekNum));
        registry.Add(new FunctionInfo("ISOWEEKNUM", 1, 1, Values, IsoWeekNum));
        registry.Add(new FunctionInfo("NETWORKDAYS", 2, 3, withHolidays, call => NetworkDays(call, weekendAt: -1, holidaysAt: 2)));
        registry.Add(new FunctionInfo("NETWORKDAYS.INTL", 2, 4, withWeekendAndHolidays, call => NetworkDays(call, weekendAt: 2, holidaysAt: 3)));
        registry.Add(new FunctionInfo("WORKDAY", 2, 3, withHolidays, call => Workday(call, weekendAt: -1, holidaysAt: 2)));
        registry.Add(new FunctionInfo("WORKDAY.INTL", 2, 4, withWeekendAndHolidays, call => Workday(call, weekendAt: 2, holidaysAt: 3)));
    }

    // Return type 1: Sunday 1 to Saturday 7; 2: Monday 1 to Sunday 7; 3: Monday 0 to Sunday 6;
    // 11-17: 1 for Monday, Tuesday ... Sunday respectively.
    private static Operand Weekday(FunctionCall call)
    {
        var serial = Serial(call, 0, logicals: true);
        if (serial.IsError)
            return serial;
        var typeValue = call.Integer(1, 1);
        if (typeValue.IsError)
            return typeValue;

        var weekday = DateSerial.MondayBasedWeekday(serial.AsNumber(), call.Context.DateSystem);
        return typeValue.AsNumber() switch
        {
            1 => CellValue.Number((weekday + 1) % 7 + 1),
            2 => CellValue.Number(weekday + 1),
            3 => CellValue.Number(weekday),
            >= 11 and <= 17 and var type => CellValue.Number((weekday - (int)(type - 11) + 7) % 7 + 1),
            _ => CellValue.Error(ErrorKind.Num),
        };
    }

    // Week 1 holds January 1; weeks start on Sunday (type 1), Monday (2) or the day of types 11-17
    // (Monday to Sunday). Type 21 is the ISO week.
    private static Operand WeekNum(FunctionCall call)
    {
        var serialValue = Serial(call, 0, logicals: true);
        if (serialValue.IsError)
            return serialValue;
        var typeValue = call.Integer(1, 1);
        if (typeValue.IsError)
            return typeValue;

        int firstDay;
        switch (typeValue.AsNumber())
        {
            case 1:
                firstDay = 6;
                break;
            case 2:
                firstDay = 0;
                break;
            case >= 11 and <= 17:
                firstDay = (int)typeValue.AsNumber() - 11;
                break;
            case 21:
                return CellValue.Number(IsoWeek(serialValue.AsNumber(), call.Context.DateSystem));
            default:
                return CellValue.Error(ErrorKind.Num);
        }

        var system = call.Context.DateSystem;
        var serial = Math.Floor(serialValue.AsNumber());
        DateSerial.TryToDate(serial, system, out var year, out _, out _);
        DateSerial.TryFromDate(year, 1, 1, system, out var january1);
        var offset = (DateSerial.MondayBasedWeekday(january1, system) - firstDay + 7) % 7;
        return CellValue.Number(Math.Floor((serial - january1 + offset) / 7) + 1);
    }

    private static Operand IsoWeekNum(FunctionCall call)
    {
        var serial = Serial(call, 0, logicals: true);
        return serial.IsError ? serial : CellValue.Number(IsoWeek(serial.AsNumber(), call.Context.DateSystem));
    }

    // The ISO week is the week of its Thursday, counted in the Thursday's year.
    private static int IsoWeek(double serial, DateSystem system)
    {
        var day = Math.Floor(serial);
        var thursday = day - DateSerial.MondayBasedWeekday(day, system) + 3;

        // Before the first day of the calendar: in 1899 (1900 system, where serial 0 stands for
        // 1899-12-31) or in 1903 (1904 system). Both years have 365 days.
        var lastDayOfPreviousYear = system == DateSystem.Date1904 ? -1 : 0;
        if (thursday <= lastDayOfPreviousYear)
            return (int)((364 + thursday - lastDayOfPreviousYear) / 7) + 1;

        DateSerial.TryToDate(thursday, system, out var year, out _, out _);
        DateSerial.TryFromDate(year, 1, 1, system, out var january1);
        return (int)((thursday - january1) / 7) + 1;
    }

    // Working days from start to end, both included; negative when the end comes first.
    private static Operand NetworkDays(FunctionCall call, int weekendAt, int holidaysAt)
    {
        var startValue = Serial(call, 0, logicals: false);
        if (startValue.IsError)
            return startValue;
        var endValue = Serial(call, 1, logicals: false);
        if (endValue.IsError)
            return endValue;
        var weekend = DefaultWeekend;
        if (weekendAt >= 0 && !TryWeekend(call, weekendAt, out weekend, out var weekendError))
            return weekendError;
        if (!TryHolidays(call, holidaysAt, out var holidays, out var holidaysError))
            return holidaysError;

        var system = call.Context.DateSystem;
        var start = (long)startValue.AsNumber();
        var end = (long)endValue.AsNumber();
        var sign = 1;
        if (start > end)
        {
            (start, end) = (end, start);
            sign = -1;
        }

        var count = CountWorkdays(start, end, weekend, system);
        foreach (var holiday in holidays)
        {
            if (holiday >= start && holiday <= end && !IsWeekend(holiday, weekend, system))
                count--;
        }

        return CellValue.Number(sign * count);
    }

    // The date a number of working days (truncated toward minus infinity) away from the start.
    private static Operand Workday(FunctionCall call, int weekendAt, int holidaysAt)
    {
        var startValue = Serial(call, 0, logicals: false);
        if (startValue.IsError)
            return startValue;
        var daysValue = Number(call, 1, logicals: false);
        if (daysValue.IsError)
            return daysValue;
        var weekend = DefaultWeekend;
        if (weekendAt >= 0 && !TryWeekend(call, weekendAt, out weekend, out var weekendError))
            return weekendError;
        if (weekend == WholeWeek)
            return CellValue.Error(ErrorKind.Value);
        if (!TryHolidays(call, holidaysAt, out var holidayList, out var holidaysError))
            return holidaysError;

        var system = call.Context.DateSystem;
        var max = (long)DateSerial.MaxSerial(system);
        var days = Math.Floor(daysValue.AsNumber());
        if (Math.Abs(days) > max)
            return CellValue.Error(ErrorKind.Num);

        var current = (long)startValue.AsNumber();
        var remaining = (long)days;
        if (remaining == 0)
            return CellValue.Number(current);

        // Step over the weekends arithmetically, then once more for each working-day holiday passed
        // on the way, until a step passes none.
        var holidays = new List<long>(holidayList);
        holidays.Sort();
        while (true)
        {
            call.Context.CancellationToken.ThrowIfCancellationRequested();
            var target = AdvanceWorkdays(current, remaining, weekend, system);
            if (target < 0 || target > max)
                return CellValue.Error(ErrorKind.Num);

            var (low, high) = remaining > 0 ? (current + 1, target) : (target, current - 1);
            var passed = 0L;
            for (var i = LowerBound(holidays, low); i < holidays.Count && holidays[i] <= high; i++)
            {
                if (!IsWeekend(holidays[i], weekend, system))
                    passed++;
            }

            if (passed == 0)
                return CellValue.Number(target);
            current = target;
            remaining = remaining > 0 ? passed : -passed;
        }
    }

    private static long CountWorkdays(long start, long end, int weekend, DateSystem system)
    {
        var length = end - start + 1;
        var count = length / 7 * (7 - CountBits(weekend));
        for (var day = start + length / 7 * 7; day <= end; day++)
        {
            if (!IsWeekend(day, weekend, system))
                count++;
        }

        return count;
    }

    // Moves a non-zero number of working days, whole weeks at a time.
    private static long AdvanceWorkdays(long start, long days, int weekend, DateSystem system)
    {
        var perWeek = 7 - CountBits(weekend);
        var step = days > 0 ? 1 : -1;
        var count = Math.Abs(days);
        var weeks = (count - 1) / perWeek;
        var left = count - weeks * perWeek;
        var day = start + step * weeks * 7;
        while (left > 0)
        {
            day += step;
            if (!IsWeekend(day, weekend, system))
                left--;
        }

        return day;
    }

    private static bool IsWeekend(long day, int weekend, DateSystem system) =>
        (weekend & (1 << DateSerial.MondayBasedWeekday(day, system))) != 0;

    private static int CountBits(int mask)
    {
        var count = 0;
        for (; mask != 0; mask &= mask - 1)
            count++;
        return count;
    }

    private static int LowerBound(List<long> sorted, long value)
    {
        var index = sorted.BinarySearch(value);
        if (index < 0)
            return ~index;
        while (index > 0 && sorted[index - 1] == value)
            index--;
        return index;
    }

    // A weekend number (1-7: two days from Saturday-Sunday to Friday-Saturday; 11-17: one day from
    // Sunday to Saturday) or seven characters 0/1 from Monday to Sunday, 1 a weekend day. Omitted is 1.
    private static bool TryWeekend(FunctionCall call, int index, out int weekend, out CellValue error)
    {
        weekend = DefaultWeekend;
        error = default;
        if (!call.Has(index))
            return true;

        var value = call.Value(index);
        if (value.Kind == CellValueKind.Text)
        {
            var text = value.AsText();
            weekend = 0;
            if (text.Length != 7)
            {
                error = CellValue.Error(ErrorKind.Value);
                return false;
            }

            for (var i = 0; i < 7; i++)
            {
                if (text[i] is not ('0' or '1'))
                {
                    error = CellValue.Error(ErrorKind.Value);
                    return false;
                }

                if (text[i] == '1')
                    weekend |= 1 << i;
            }

            return true;
        }

        var number = Number(call, index, logicals: false);
        if (number.IsError)
        {
            error = number;
            return false;
        }

        switch (Math.Truncate(number.AsNumber()))
        {
            case >= 1 and <= 7 and var code:
                weekend = (1 << (int)((code + 4) % 7)) | (1 << (int)((code + 5) % 7));
                return true;
            case >= 11 and <= 17 and var code:
                weekend = 1 << (int)((code - 5) % 7);
                return true;
            default:
                error = CellValue.Error(ErrorKind.Num);
                return false;
        }
    }

    // Holiday dates from a range, an array or a single value; empty cells are skipped, the first
    // error wins, TRUE/FALSE and non-date text are #VALUE!, a negative date is #NUM!.
    private static bool TryHolidays(FunctionCall call, int index, out HashSet<long> holidays, out CellValue error)
    {
        holidays = [];
        error = default;
        if (!call.Has(index))
            return true;

        var context = call.Context;
        var max = DateSerial.MaxSerial(context.DateSystem);
        var found = holidays;
        CellValue? failure = null;
        Aggregation.ForEachIn(call, index, (value, _) =>
        {
            if (value.Kind is CellValueKind.Empty or CellValueKind.Missing)
                return true;
            if (value.Kind == CellValueKind.Boolean)
            {
                failure = CellValue.Error(ErrorKind.Value);
                return false;
            }

            var number = Coercion.ToNumber(value, context.Culture, context.DateSystem);
            if (number.IsError)
            {
                failure = number;
                return false;
            }

            var serial = number.AsNumber();
            if (!(serial >= 0) || serial >= max + 1)
            {
                failure = CellValue.Error(ErrorKind.Num);
                return false;
            }

            found.Add((long)serial);
            return true;
        });

        if (failure is { } failed)
        {
            error = failed;
            return false;
        }

        return true;
    }
}
