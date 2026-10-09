using System;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// Date and time functions. Dates are serial numbers (<see cref="DateSerial"/>); the calendar
/// functions are here, spans in DateTimeSpans.cs, weeks and workdays in DateTimeWeeks.cs.
/// </summary>
internal static partial class DateTimeFunctions
{
    private static readonly ArgumentKind[] Values = [ArgumentKind.Value];

    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("NOW", 0, 0, Values, Now) { IsVolatile = true });
        registry.Add(new FunctionInfo("TODAY", 0, 0, Values, Today) { IsVolatile = true });
        registry.Add(new FunctionInfo("DATE", 3, 3, Values, Date));
        registry.Add(new FunctionInfo("TIME", 3, 3, Values, Time));
        registry.Add(new FunctionInfo("YEAR", 1, 1, Values, call => DatePart(call, (year, _, _) => year)));
        registry.Add(new FunctionInfo("MONTH", 1, 1, Values, call => DatePart(call, (_, month, _) => month)));
        registry.Add(new FunctionInfo("DAY", 1, 1, Values, call => DatePart(call, (_, _, day) => day)));
        registry.Add(new FunctionInfo("HOUR", 1, 1, Values, call => TimePart(call, time => time.Hour)));
        registry.Add(new FunctionInfo("MINUTE", 1, 1, Values, call => TimePart(call, time => time.Minute)));
        registry.Add(new FunctionInfo("SECOND", 1, 1, Values, call => TimePart(call, time => time.Second)));
        registry.Add(new FunctionInfo("DATEVALUE", 1, 1, Values, call => TextValue(call, serial => Math.Floor(serial))));
        registry.Add(new FunctionInfo("TIMEVALUE", 1, 1, Values, call => TextValue(call, serial => serial - Math.Floor(serial))));
        RegisterSpans(registry);
        RegisterWeeks(registry);
    }

    private static Operand Now(FunctionCall call)
    {
        var workbook = call.Context.Workbook;
        return CellValue.Number(DateSerial.FromDateTime(workbook.Clock.GetLocalNow().DateTime, workbook.DateSystem));
    }

    private static Operand Today(FunctionCall call)
    {
        var workbook = call.Context.Workbook;
        return CellValue.Number(DateSerial.FromDateTime(workbook.Clock.GetLocalNow().DateTime.Date, workbook.DateSystem));
    }

    // Arguments are truncated; months and days carry over (DATE(2020,13,1) is 2021-01-01), years
    // 0-1899 count from 1900. Outside 1900 (1904) to 9999 the result is #NUM!.
    private static Operand Date(FunctionCall call)
    {
        Span<long> parts = stackalloc long[3];
        for (var i = 0; i < 3; i++)
        {
            var part = call.Integer(i);
            if (part.IsError)
                return part;

            // Far beyond any carry-over that can land in the calendar; also keeps the cast exact.
            if (Math.Abs(part.AsNumber()) > 1e8)
                return CellValue.Error(ErrorKind.Num);
            parts[i] = (long)part.AsNumber();
        }

        return DateSerial.TryFromDate(parts[0], parts[1], parts[2], call.Context.DateSystem, out var serial)
            ? CellValue.Number(serial)
            : CellValue.Error(ErrorKind.Num);
    }

    // The fraction of a day; whole days are dropped (TIME(25,0,0) is 1:00). Arguments are truncated,
    // each at most 32767, and may be negative while the total is not.
    private static Operand Time(FunctionCall call)
    {
        var seconds = 0.0;
        Span<double> scale = [3600, 60, 1];
        for (var i = 0; i < 3; i++)
        {
            var part = call.Integer(i);
            if (part.IsError)
                return part;
            if (part.AsNumber() > 32767)
                return CellValue.Error(ErrorKind.Num);
            seconds += part.AsNumber() * scale[i];
        }

        if (seconds < 0)
            return CellValue.Error(ErrorKind.Num);
        var days = seconds / 86400;
        return CellValue.Number(days - Math.Floor(days));
    }

    private static Operand DatePart(FunctionCall call, Func<int, int, int, int> pick)
    {
        var serial = call.Number(0);
        if (serial.IsError)
            return serial;
        return DateSerial.TryToDate(serial.AsNumber(), call.Context.DateSystem, out var year, out var month, out var day)
            ? CellValue.Number(pick(year, month, day))
            : CellValue.Error(ErrorKind.Num);
    }

    private static Operand TimePart(FunctionCall call, Func<(int Hour, int Minute, int Second), int> pick)
    {
        var serial = Serial(call, 0, logicals: true);
        return serial.IsError ? serial : CellValue.Number(pick(DateSerial.TimeOfDay(serial.AsNumber())));
    }

    // DATEVALUE and TIMEVALUE read text only: a number, even a date, is #VALUE!. Text without a year
    // ("1-Jan") is in the current year by the workbook clock.
    private static Operand TextValue(FunctionCall call, Func<double, double> part)
    {
        var value = call.Value(0);
        if (value.IsError)
            return value;
        if (value.Kind != CellValueKind.Text)
            return CellValue.Error(ErrorKind.Value);

        var context = call.Context;
        var year = context.Workbook.Clock.GetLocalNow().Year;
        return DateText.TryParse(value.AsText().Trim(), context.Culture, context.DateSystem, year, out var serial)
            ? CellValue.Number(part(serial))
            : CellValue.Error(ErrorKind.Value);
    }

    /// <summary>
    /// A date argument as a serial number: a number from 0 to the last day of 9999, else #NUM!.
    /// The Analysis ToolPak functions (EDATE, WORKDAY, YEARFRAC...) refuse TRUE and FALSE with
    /// #VALUE!; the others read them as 1 and 0.
    /// </summary>
    private static CellValue Serial(FunctionCall call, int index, bool logicals)
    {
        var number = Number(call, index, logicals);
        if (number.IsError)
            return number;
        var serial = number.AsNumber();
        return serial >= 0 && serial < DateSerial.MaxSerial(call.Context.DateSystem) + 1
            ? number
            : CellValue.Error(ErrorKind.Num);
    }

    private static CellValue Number(FunctionCall call, int index, bool logicals)
    {
        if (!logicals && call.Value(index).Kind == CellValueKind.Boolean)
            return CellValue.Error(ErrorKind.Value);
        return call.Number(index);
    }
}
