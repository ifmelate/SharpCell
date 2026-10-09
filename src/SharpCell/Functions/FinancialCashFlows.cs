using System;
using System.Collections.Generic;
using SharpCell.Evaluation;
using static SharpCell.Functions.FinancialArguments;

namespace SharpCell.Functions;

/// <summary>Series of cash flows: NPV, IRR, MIRR, XNPV, XIRR and FVSCHEDULE.</summary>
internal static class FinancialCashFlows
{
    private const int CancellationCheckInterval = 4096;

    // IRR and XIRR: Newton's method from the guess (Excel's default 0.1), done when a step is
    // below the tolerance. Plain Newton overshoots past −100% on many series Excel solves, so when
    // it fails the root is bracketed and bisected instead.
    private const int NewtonIterations = 100;
    private const int BisectionIterations = 50;

    public static void Register(FunctionRegistry registry)
    {
        var max = FunctionRegistry.MaxArguments;
        registry.Add(new FunctionInfo("NPV", 2, max, [ArgumentKind.Value, ArgumentKind.Any], NetPresentValue));
        registry.Add(new FunctionInfo("IRR", 1, 2, [ArgumentKind.ArrayContext, ArgumentKind.Value], InternalRate));
        registry.Add(new FunctionInfo("MIRR", 3, 3, [ArgumentKind.ArrayContext, ArgumentKind.Value], ModifiedInternalRate));
        registry.Add(new FunctionInfo("XNPV", 3, 3, [ArgumentKind.Value, ArgumentKind.ArrayContext], ScheduledPresentValue));
        registry.Add(new FunctionInfo("XIRR", 2, 3, [ArgumentKind.ArrayContext, ArgumentKind.ArrayContext, ArgumentKind.Value], ScheduledInternalRate));
        registry.Add(new FunctionInfo("FVSCHEDULE", 2, 2, [ArgumentKind.Value, ArgumentKind.ArrayContext], FutureValueOfSchedule));
    }

    // NPV(rate, value1, ...): the values are discounted from period 1. Numbers in ranges and
    // arrays count, other values there are skipped; typed values convert as in SUM.
    private static Operand NetPresentValue(FunctionCall call)
    {
        var rate = call.Number(0);
        if (rate.IsError)
            return rate;

        var values = new List<double>();
        var error = Numbers(call, 1, call.Count, values);
        if (error.IsError)
            return error;
        return CellValue.Number(Npv(rate.AsNumber(), values));
    }

    private static double Npv(double rate, List<double> values)
    {
        var total = 0.0;
        for (var i = 0; i < values.Count; i++)
            total += values[i] / Math.Pow(1 + rate, i + 1);
        return total;
    }

    private static double NpvSlope(double rate, List<double> values)
    {
        var total = 0.0;
        for (var i = 0; i < values.Count; i++)
            total -= values[i] * (i + 1) / Math.Pow(1 + rate, i + 2);
        return total;
    }

    // IRR(values, [guess])
    private static Operand InternalRate(FunctionCall call)
    {
        var values = new List<double>();
        var error = Numbers(call, 0, 1, values);
        if (error.IsError)
            return error;
        var guess = call.Number(1, 0.1);
        if (guess.IsError)
            return guess;
        if (guess.AsNumber() <= -1)
            return CellValue.Error(ErrorKind.Value);
        if (!HasBothSigns(values))
            return CellValue.Error(ErrorKind.Num);

        var result = Solve(r => Npv(r, values), r => NpvSlope(r, values), guess.AsNumber(), 1e-8, 1e-10, -0.99999, tryBelow: true);
        return result is { } r ? CellValue.Number(r) : CellValue.Error(ErrorKind.Num);
    }

    // MIRR(values, finance_rate, reinvest_rate): negative flows are financed at one rate, positive
    // ones reinvested at the other: (−NPV(reinvest, positive)·(1+reinvest)^n / (NPV(finance, negative)·(1+finance)))^(1/(n−1)) − 1.
    private static Operand ModifiedInternalRate(FunctionCall call)
    {
        var values = new List<double>();
        var error = Numbers(call, 0, 1, values);
        if (error.IsError)
            return error;
        var finance = call.Number(1);
        if (finance.IsError)
            return finance;
        var reinvest = call.Number(2);
        if (reinvest.IsError)
            return reinvest;
        var (financeRate, reinvestRate) = (finance.AsNumber(), reinvest.AsNumber());

        var positive = new List<double>(values.Count);
        var negative = new List<double>(values.Count);
        var lastNegative = -1;
        for (var i = 0; i < values.Count; i++)
        {
            positive.Add(values[i] >= 0 ? values[i] : 0);
            negative.Add(values[i] < 0 ? values[i] : 0);
            if (values[i] < 0)
                lastNegative = i;
        }

        if (lastNegative < 0)
            return CellValue.Error(ErrorKind.Div0);

        // At a rate of −100% the discounting divides by zero; the limits are finite or irrelevant.
        var n = values.Count;
        var top = reinvestRate == -1
            ? positive[^1]
            : -Npv(reinvestRate, positive) * Math.Pow(1 + reinvestRate, n);
        var bottom = financeRate == -1
            ? lastNegative == 0 ? negative[0] : double.PositiveInfinity
            : Npv(financeRate, negative) * (1 + financeRate);
        var result = Math.Pow(top / bottom, 1.0 / (n - 1)) - 1;
        return double.IsInfinity(result) ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(result);
    }

    // XNPV(rate, values, dates): Σ value / (1+rate)^((date − first date)/365).
    private static Operand ScheduledPresentValue(FunctionCall call)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required))
            return error;
        var rate = x[0];
        if (!TrySchedule(call, 1, 2, ErrorKind.Num, emptyIsZero: false, out var values, out var dates, out error))
            return error;
        if (rate <= 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(Xnpv(rate, values, dates));
    }

    private static double Xnpv(double rate, double[] values, double[] dates)
    {
        var total = values[0];
        for (var i = 1; i < values.Length; i++)
            total += values[i] / Math.Pow(1 + rate, (dates[i] - dates[0]) / 365);
        return total;
    }

    private static double XnpvSlope(double rate, double[] values, double[] dates)
    {
        var total = 0.0;
        for (var i = 1; i < values.Length; i++)
        {
            var years = (dates[i] - dates[0]) / 365;
            total -= values[i] * years / Math.Pow(1 + rate, years + 1);
        }

        return total;
    }

    // XIRR(values, dates, [guess])
    private static Operand ScheduledInternalRate(FunctionCall call)
    {
        if (!TrySchedule(call, 0, 1, ErrorKind.Value, emptyIsZero: true, out var values, out var dates, out var error))
            return error;
        if (!TryAddInNumber(call, 2, 0.1, out var guess, out error))
            return error;
        if (guess <= -1)
            return CellValue.Error(ErrorKind.Value);
        if (!HasBothSigns(values))
            return CellValue.Error(ErrorKind.Num);

        var result = Solve(r => Xnpv(r, values, dates), r => XnpvSlope(r, values, dates), guess, 1e-7, 1e-8, -0.9999, tryBelow: false);
        return result is { } r ? CellValue.Number(r) : CellValue.Error(ErrorKind.Num);
    }

    // FVSCHEDULE(principal, schedule): the principal grown by each rate in turn; empty cells count as 0.
    private static Operand FutureValueOfSchedule(FunctionCall call)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required))
            return error;
        var result = x[0];
        var schedule = call.Value(1);
        if (schedule.Kind != CellValueKind.Array)
            schedule = CellValue.Array(new[,] { { schedule } });
        foreach (var element in schedule.AsArray())
        {
            if (element.Kind is CellValueKind.Empty or CellValueKind.Missing)
                continue;
            if (element.Kind == CellValueKind.Boolean)
                return CellValue.Error(ErrorKind.Value);
            var rate = Coercion.ToNumber(element, call.Context.Culture, call.Context.DateSystem);
            if (rate.IsError)
                return rate;
            result *= 1 + rate.AsNumber();
        }

        return CellValue.Number(result);
    }

    private static bool HasBothSigns(IReadOnlyList<double> values)
    {
        bool positive = false, negative = false;
        foreach (var value in values)
        {
            positive |= value > 0;
            negative |= value < 0;
        }

        return positive && negative;
    }

    /// <summary>
    /// A root of <paramref name="f"/>: Newton from the guess; failing that, bisection on
    /// [<paramref name="low"/>, 100] when the ends bracket a root, else Newton from far above
    /// (and, with <paramref name="tryBelow"/>, from below −100%).
    /// </summary>
    private static double? Solve(Func<double, double> f, Func<double, double> slope, double guess, double newtonTolerance,
        double bisectionTolerance, double low, bool tryBelow)
    {
        if (Newton(f, slope, guess, newtonTolerance) is { } root)
            return root;

        const double high = 100;
        var fLow = f(low);
        var fHigh = f(high);
        if (!double.IsFinite(fLow) || !double.IsFinite(fHigh))
            return null;
        if (fLow * fHigh > 0)
            return Newton(f, slope, 200, newtonTolerance) ?? (tryBelow ? Newton(f, slope, -2, newtonTolerance) : null);

        // Keep the end where f is negative and halve the step towards the other end.
        var (at, step) = fLow < 0 ? (low, high - low) : (high, low - high);
        for (var i = 1; i < BisectionIterations; i++)
        {
            step *= 0.5;
            var middle = at + step;
            var value = f(middle);
            if (value <= 0)
                at = middle;
            if (Math.Abs(value) < bisectionTolerance || Math.Abs(step) < bisectionTolerance)
                return middle;
        }

        return null;
    }

    private static double? Newton(Func<double, double> f, Func<double, double> slope, double guess, double tolerance)
    {
        var x = guess;
        for (var i = 0; i < NewtonIterations; i++)
        {
            var next = x - f(x) / slope(x);
            if (!double.IsFinite(next))
                return null;
            // A root at or below −100% is not a rate of return.
            if (Math.Abs(next - x) < tolerance)
                return next > -1 ? next : null;
            x = next;
        }

        return null;
    }

    /// <summary>
    /// The numbers of arguments <paramref name="first"/> to <paramref name="end"/> (exclusive): numbers
    /// in references and arrays, other values there skipped; typed values converted. Every cell is
    /// read even after an error, so all dirty inputs are found in one pass. Returns the first error.
    /// </summary>
    private static CellValue Numbers(FunctionCall call, int first, int end, List<double> numbers)
    {
        var context = call.Context;
        CellValue? error = null;
        var visited = 0;
        for (var i = first; i < end; i++)
        {
            var argument = call[i];
            if (argument.Reference is { } reference)
            {
                foreach (var (sheet, area) in reference.Areas)
                {
                    foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
                    {
                        if (++visited % CancellationCheckInterval == 0)
                            context.CancellationToken.ThrowIfCancellationRequested();
                        Add(context.ReadCell(sheet, cell), direct: false);
                    }
                }
            }
            else if (argument.Value.Kind == CellValueKind.Array)
            {
                foreach (var element in argument.Value.AsArray())
                    Add(element, direct: false);
            }
            else
            {
                Add(argument.Value.Kind == CellValueKind.Missing ? CellValue.Number(0) : argument.Value, direct: true);
            }
        }

        return error ?? CellValue.Empty;

        void Add(CellValue value, bool direct)
        {
            if (error is not null)
                return;
            if (value.IsError)
                error = value;
            else if (value.Kind == CellValueKind.Number)
                numbers.Add(value.AsNumber());
            else if (direct && value.Kind != CellValueKind.Empty)
            {
                var number = Coercion.ToNumber(value, context.Culture, context.DateSystem);
                if (number.IsError)
                    error = number;
                else
                    numbers.Add(number.AsNumber());
            }
        }
    }

    /// <summary>
    /// The values and dates of XNPV and XIRR, paired by position. A value that is not a number gives
    /// <paramref name="notNumber"/> (an empty one counts as 0 for XIRR); dates are truncated to whole
    /// days, must be valid and not before the first.
    /// </summary>
    private static bool TrySchedule(FunctionCall call, int valuesIndex, int datesIndex, ErrorKind notNumber, bool emptyIsZero,
        out double[] values, out double[] dates, out CellValue error)
    {
        values = dates = [];
        if (!TryFlatten(call.Value(valuesIndex), notNumber, emptyIsZero, out values, out error)
            || !TryFlatten(call.Value(datesIndex), notNumber, emptyIsZero, out dates, out error))
            return false;

        error = CellValue.Error(ErrorKind.Num);
        if (values.Length != dates.Length || values.Length == 0)
            return false;

        var maxSerial = DateSerial.MaxSerial(call.Context.DateSystem);
        for (var i = 0; i < dates.Length; i++)
        {
            dates[i] = Math.Floor(dates[i]);
            if (dates[i] < 0 || dates[i] > maxSerial || dates[i] < dates[0])
                return false;
        }

        error = default;
        return true;
    }

    private static bool TryFlatten(CellValue value, ErrorKind notNumber, bool emptyIsZero, out double[] numbers, out CellValue error)
    {
        numbers = [];
        var elements = value.Kind == CellValueKind.Array ? value.AsArray() : new[,] { { value } };
        var result = new double[elements.Length];
        var i = 0;
        foreach (var element in elements)
        {
            if (element.IsError)
            {
                error = element;
                return false;
            }

            if (element.Kind == CellValueKind.Number)
                result[i++] = element.AsNumber();
            else if (emptyIsZero && element.Kind is CellValueKind.Empty or CellValueKind.Missing)
                result[i++] = 0;
            else
            {
                error = CellValue.Error(element.Kind is CellValueKind.Empty or CellValueKind.Missing ? ErrorKind.Num : notNumber);
                return false;
            }
        }

        numbers = result;
        error = default;
        return true;
    }
}
