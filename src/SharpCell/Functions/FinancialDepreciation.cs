using System;
using SharpCell.Evaluation;
using static SharpCell.Functions.FinancialArguments;

namespace SharpCell.Functions;

/// <summary>Depreciation: SLN, SYD, DB, DDB, VDB and the French AMORLINC and AMORDEGRC.</summary>
internal static class FinancialDepreciation
{
    private static readonly ArgumentKind[] Scalars = [ArgumentKind.Value];

    private const int CancellationCheckInterval = 4096;

    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("SLN", 3, 3, Scalars, StraightLine));
        registry.Add(new FunctionInfo("SYD", 4, 4, Scalars, SumOfYearsDigits));
        registry.Add(new FunctionInfo("DB", 4, 5, Scalars, FixedDecliningBalance));
        registry.Add(new FunctionInfo("DDB", 4, 5, Scalars, DoubleDecliningBalance));
        registry.Add(new FunctionInfo("VDB", 5, 7, Scalars, VariableDecliningBalance));
        registry.Add(new FunctionInfo("AMORLINC", 6, 7, Scalars, call => French(call, degressive: false)));
        registry.Add(new FunctionInfo("AMORDEGRC", 6, 7, Scalars, call => French(call, degressive: true)));
    }

    // SLN(cost, salvage, life)
    private static Operand StraightLine(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required))
            return error;
        var (cost, salvage, life) = (x[0], x[1], x[2]);
        return life == 0 ? CellValue.Error(ErrorKind.Div0) : CellValue.Number((cost - salvage) / life);
    }

    // SYD(cost, salvage, life, per)
    private static Operand SumOfYearsDigits(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required, Required))
            return error;
        var (cost, salvage, life, period) = (x[0], x[1], x[2], x[3]);
        if (life == 0 || period > life || period <= 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number((cost - salvage) * (life - period + 1) * 2 / (life * (life + 1)));
    }

    // DB(cost, salvage, life, period, [month]): the rate 1 − (salvage/cost)^(1/life), rounded to
    // three decimals, on the declining balance; the first year has only `month` months, and when
    // that is less than 12 the year after the life takes the rest.
    private static Operand FixedDecliningBalance(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required, Required, 12))
            return error;
        var (cost, salvage, life, period, month) = (x[0], x[1], x[2], x[3], Math.Truncate(x[4]));
        if ((month == 12 && period > life) || period > life + 1 || month <= 0 || month > 12 || period <= 0 || cost < 0)
            return CellValue.Error(ErrorKind.Num);
        if (cost == 0)
            return CellValue.Number(0);

        var rate = Math.Round((1 - Math.Pow(salvage / cost, 1 / life)) * 1000, MidpointRounding.AwayFromZero) / 1000;
        var whole = (long)Math.Truncate(period);
        var depreciated = cost * rate * month / 12;
        if (whole <= 1)
            return CellValue.Number(depreciated);
        for (long p = 2; p < whole; p++)
        {
            if (p % CancellationCheckInterval == 0)
                call.Context.CancellationToken.ThrowIfCancellationRequested();
            depreciated += (cost - depreciated) * rate;
        }

        var result = whole == (long)Math.Truncate(life) + 1
            ? (cost - depreciated) * rate * (12 - month) / 12
            : (cost - depreciated) * rate;
        return CellValue.Number(result);
    }

    // DDB(cost, salvage, life, period, [factor]): factor/life of the remaining value each period,
    // never below salvage.
    private static Operand DoubleDecliningBalance(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required, Required, 2))
            return error;
        var (cost, salvage, life, period, factor) = (x[0], x[1], x[2], x[3], x[4]);
        if (period > life || cost < 0 || salvage < 0 || period <= 0 || factor <= 0)
            return CellValue.Error(ErrorKind.Num);

        // A fractional first period counts as the whole first period.
        period = Math.Max(period, 1);
        var rate = factor / life;
        double before;
        if (rate >= 1)
        {
            rate = 1;
            before = period == 1 ? cost : 0;
        }
        else
        {
            before = cost * Math.Pow(1 - rate, period - 1);
        }

        var after = cost * Math.Pow(1 - rate, period);
        var result = after < salvage ? before - salvage : before - after;
        return CellValue.Number(Math.Max(result, 0));
    }

    // VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch]): declining balance
    // between two (fractional) periods, switching to straight line when that is larger unless no_switch.
    private static Operand VariableDecliningBalance(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required, Required, Required, 2, 0))
            return error;
        var (cost, salvage, life, first, last, factor) = (x[0], x[1], x[2], x[3], x[4], x[5]);
        var straightLine = x[6] == 0;
        if (cost < 0 || salvage < 0 || life <= 0 || factor <= 0 || first < 0 || first > life || last > life
            || first > last || last <= 0)
            return CellValue.Error(ErrorKind.Num);
        if (straightLine && life == first && first == last)
            return CellValue.Error(ErrorKind.Num);

        var cancellation = call.Context.CancellationToken;
        var result = TotalDepreciation(cost, salvage, life, last, factor, straightLine, cancellation)
                     - TotalDepreciation(cost, salvage, life, first, factor, straightLine, cancellation);
        return CellValue.Number(result);
    }

    // Depreciation from the start to `period`, period by period, the last one prorated.
    private static double TotalDepreciation(double cost, double salvage, double life, double period, double factor,
        bool straightLine, System.Threading.CancellationToken cancellation)
    {
        var fraction = period - Math.Truncate(period);
        var whole = (long)Math.Truncate(period);
        var wholeLife = (long)Math.Truncate(life);
        var total = 0.0;
        for (long p = 0; ; p++)
        {
            if (p % CancellationCheckInterval == CancellationCheckInterval - 1)
                cancellation.ThrowIfCancellationRequested();
            var next = total + Depreciation(cost, salvage, life, factor, straightLine, total, p);
            if (whole == 0)
                return next * fraction;
            if (p == whole - 1)
            {
                var following = straightLine && DecliningBalance(cost, salvage, life, factor, next) < StraightLine(cost, salvage, life, next, p + 1)
                    ? whole == wholeLife ? 0 : StraightLine(cost, salvage, life, next, p + 1)
                    : DecliningBalance(cost, salvage, life, factor, next);
                return next + following * fraction;
            }

            total = next;
        }
    }

    private static double Depreciation(double cost, double salvage, double life, double factor, bool straightLine, double total, long period)
    {
        var declining = DecliningBalance(cost, salvage, life, factor, total);
        var straight = StraightLine(cost, salvage, life, total, period);
        return straightLine && declining < straight ? straight : declining;
    }

    private static double DecliningBalance(double cost, double salvage, double life, double factor, double total) =>
        Math.Min((cost - total) * (factor / life), cost - salvage - total);

    private static double StraightLine(double cost, double salvage, double life, double total, double period) =>
        (cost - total - salvage) / (life - period);

    // AMORLINC and AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate, [basis]): the
    // French accounting system. The first period is prorated by days; AMORDEGRC multiplies the rate
    // by a coefficient for the asset's life and rounds each period's depreciation to a whole number.
    private static Operand French(FunctionCall call, bool degressive)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required, Required, Required, Required, Required, 0))
            return error;
        var (cost, salvage, period, rate, basis) = (x[0], x[3], x[4], x[5], Math.Truncate(x[6]));
        if (!TryDate(x[1], call.Context.DateSystem, out var purchased) || !TryDate(x[2], call.Context.DateSystem, out var firstPeriod))
            return CellValue.Error(ErrorKind.Num);
        if (cost < 0 || salvage < 0 || salvage >= cost || period < 0 || rate <= 0 || purchased >= firstPeriod
            || basis is < 0 or > 4 or 2)
            return CellValue.Error(ErrorKind.Num);

        var life = 1 / rate;
        if (degressive && (life is >= 0 and <= 3 || life is >= 4 and <= 5))
            return CellValue.Error(ErrorKind.Num);

        var b = (int)basis;
        return CellValue.Number(degressive
            ? Degressive(cost, purchased, firstPeriod, salvage, period, rate, b)
            : Linear(cost, purchased, firstPeriod, salvage, period, rate, b));
    }

    private static double Linear(double cost, FinancialDate purchased, FinancialDate firstPeriod, double salvage, double period, double rate, int basis)
    {
        var life = Math.Ceiling(1 / rate);
        if (cost == salvage || period > life)
            return 0;
        var (first, _) = FirstDepreciation(cost, purchased, firstPeriod, salvage, rate, life, basis);
        if (period == 0)
            return first;

        // A constant rate·cost per period, until the depreciable amount runs out.
        var depreciation = rate * cost;
        var available = cost - salvage - first;
        for (var counted = 1.0; counted <= period; counted++)
        {
            depreciation = Math.Min(depreciation, available);
            available = Math.Max(available - depreciation, 0);
        }

        return depreciation;
    }

    private static double Degressive(double cost, FinancialDate purchased, FinancialDate firstPeriod, double salvage, double period, double rate, int basis)
    {
        var life = Math.Ceiling(1 / rate);
        if (cost == salvage || period > life)
            return 0;
        var coefficient = life switch
        {
            >= 3 and <= 4 => 1.5,
            >= 5 and <= 6 => 2.0,
            > 6 => 2.5,
            _ => 1.0,
        };
        var degressiveRate = rate * coefficient;
        var (firstLinear, assetLife) = FirstDepreciation(cost, purchased, firstPeriod, salvage, degressiveRate, life, basis);
        var first = RoundToWhole(firstLinear);
        if (period == 0)
            return first;

        // Two periods before the end the rest is split in halves.
        var depreciation = 0.0;
        var remaining = cost - first;
        for (var counted = 1.0; counted <= period;)
        {
            counted++;
            var lastTwo = Math.Abs(assetLife - counted - 2) < 0.0001;
            var amount = lastTwo ? remaining * 0.5 : degressiveRate * remaining;
            if (lastTwo)
                degressiveRate = 1;
            depreciation = remaining < salvage ? Math.Max(remaining - salvage, 0) : amount;
            remaining -= depreciation;
        }

        return RoundToWhole(depreciation);
    }

    // The prorated first period and the asset life it implies (a year longer when prorated).
    private static (double Depreciation, double Life) FirstDepreciation(double cost, FinancialDate purchased, FinancialDate firstPeriod,
        double salvage, double rate, double life, int basis)
    {
        var yearLength = basis switch
        {
            1 => FinancialDayCount.IsLeapYear(purchased.Year) ? 366.0 : 365.0,
            3 => 365.0,
            _ => 360.0,
        };
        var days = FinancialDayCount.DaysBetween(WithoutLeapDay(purchased, basis), WithoutLeapDay(firstPeriod, basis), basis);
        var prorated = days / yearLength * rate * cost;
        var first = prorated == 0 ? cost * rate : prorated;
        return (Math.Min(first, cost - salvage), prorated == 0 ? life : life + 1);
    }

    // The actual-day bases count a leap year's February as ending on the 28th.
    private static FinancialDate WithoutLeapDay(FinancialDate date, int basis) =>
        (basis is 1 or 3) && FinancialDayCount.IsLeapYear(date.Year) && date.Month == 2 && date.Day >= 28
            ? date with { Day = 28 }
            : date;

    // Rounded to 13 significant decimals first, as Excel's stored precision, then to a whole number.
    private static double RoundToWhole(double value) =>
        Math.Round(Math.Round(value * 1e13, MidpointRounding.AwayFromZero) / 1e13, MidpointRounding.AwayFromZero);
}
