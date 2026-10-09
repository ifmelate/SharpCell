using System;
using SharpCell.Evaluation;
using static SharpCell.Functions.FinancialArguments;

namespace SharpCell.Functions;

/// <summary>
/// Coupon bonds: the coupon schedule (COUP*), price and yield of regular bonds and of bonds with an
/// odd first or last period, and duration. Prices are per 100 of face value, clean (without the
/// accrued interest).
/// </summary>
internal static class FinancialBonds
{
    private static readonly ArgumentKind[] Scalars = [ArgumentKind.Value];

    // Yields are solved by Newton's method on the price, as Excel: at most 100 steps.
    private const int YieldIterations = 100;
    private const double YieldTolerance = 1e-10;

    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("COUPDAYBS", 3, 4, Scalars, call => Coupon(call, CouponDaysBefore)));
        registry.Add(new FunctionInfo("COUPDAYS", 3, 4, Scalars, call => Coupon(call, FinancialDayCount.CouponDays)));
        registry.Add(new FunctionInfo("COUPDAYSNC", 3, 4, Scalars, call => Coupon(call, FinancialDayCount.CouponDaysAfterSettlement)));
        registry.Add(new FunctionInfo("COUPNCD", 3, 4, Scalars, call => Coupon(call, (s, m, f, _) =>
            FinancialDayCount.CouponDates(s, m, f).Next.ToSerial(call.Context.DateSystem))));
        registry.Add(new FunctionInfo("COUPPCD", 3, 4, Scalars, call => Coupon(call, (s, m, f, _) =>
            FinancialDayCount.CouponDates(s, m, f).Previous.ToSerial(call.Context.DateSystem))));
        registry.Add(new FunctionInfo("COUPNUM", 3, 4, Scalars, call => Coupon(call, (s, m, f, _) => FinancialDayCount.CouponCount(s, m, f))));
        registry.Add(new FunctionInfo("PRICE", 6, 7, Scalars, Price));
        registry.Add(new FunctionInfo("YIELD", 6, 7, Scalars, Yield));
        registry.Add(new FunctionInfo("DURATION", 5, 6, Scalars, call => Duration(call, modified: false)));
        registry.Add(new FunctionInfo("MDURATION", 5, 6, Scalars, call => Duration(call, modified: true)));
        registry.Add(new FunctionInfo("ODDFPRICE", 8, 9, Scalars, call => OddFirst(call, price: true)));
        registry.Add(new FunctionInfo("ODDFYIELD", 8, 9, Scalars, call => OddFirst(call, price: false)));
        registry.Add(new FunctionInfo("ODDLPRICE", 7, 8, Scalars, call => OddLast(call, price: true)));
        registry.Add(new FunctionInfo("ODDLYIELD", 7, 8, Scalars, call => OddLast(call, price: false)));
    }

    // COUP*(settlement, maturity, frequency, [basis])
    private static Operand Coupon(FunctionCall call, Func<FinancialDate, FinancialDate, int, int, double> measure)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required, Required, 0))
            return error;
        var system = call.Context.DateSystem;
        if (!TryDate(x[0], system, out var settlement) || !TryDate(x[1], system, out var maturity)
            || !FinancialSecurities.TryFrequency(x[2], out var frequency) || !FinancialSecurities.TryBasis(x[3], out var basis)
            || settlement >= maturity)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(measure(settlement, maturity, frequency, basis));
    }

    private static double CouponDaysBefore(FinancialDate settlement, FinancialDate maturity, int frequency, int basis) =>
        FinancialDayCount.CouponDaysBeforeSettlement(settlement, maturity, frequency, basis);

    /// <summary>The regular bond's schedule seen from settlement.</summary>
    private readonly record struct Schedule(double Coupons, double PeriodDays, double AccruedDays)
    {
        public double ToNextCoupon => PeriodDays - AccruedDays;

        public static Schedule Of(FinancialDate settlement, FinancialDate maturity, int frequency, int basis) => new(
            FinancialDayCount.CouponCount(settlement, maturity, frequency),
            FinancialDayCount.CouponDays(settlement, maturity, frequency, basis),
            FinancialDayCount.CouponDaysBeforeSettlement(settlement, maturity, frequency, basis));
    }

    // Reads (settlement, maturity, rate, yld or pr, redemption, frequency, [basis]).
    private static bool TryBond(FunctionCall call, bool pricing, out FinancialDate settlement, out FinancialDate maturity,
        out double rate, out double value, out double redemption, out int frequency, out int basis, out CellValue error)
    {
        settlement = maturity = default;
        rate = value = redemption = 0;
        frequency = basis = 0;
        if (!TryAddInNumbers(call, out var x, out error, Required, Required, Required, Required, Required, Required, 0))
            return false;
        (rate, value, redemption) = (x[2], x[3], x[4]);
        var system = call.Context.DateSystem;
        error = CellValue.Error(ErrorKind.Num);
        return TryDate(x[0], system, out settlement) && TryDate(x[1], system, out maturity)
               && FinancialSecurities.TryFrequency(x[5], out frequency) && FinancialSecurities.TryBasis(x[6], out basis)
               && rate >= 0 && (pricing ? value >= 0 : value > 0) && redemption > 0 && settlement < maturity;
    }

    // PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])
    private static Operand Price(FunctionCall call)
    {
        if (!TryBond(call, pricing: true, out var settlement, out var maturity, out var rate, out var yield, out var redemption,
                out var frequency, out var basis, out var error))
            return error;
        var schedule = Schedule.Of(settlement, maturity, frequency, basis);
        return CellValue.Number(BondPrice(schedule, rate, yield, redemption, frequency));
    }

    private static double BondPrice(Schedule schedule, double rate, double yield, double redemption, int frequency)
    {
        var coupon = 100 * rate / frequency;
        var accrued = coupon * schedule.AccruedDays / schedule.PeriodDays;
        var toNext = schedule.ToNextCoupon / schedule.PeriodDays;

        // With one coupon left the discounting is simple interest.
        if (schedule.Coupons == 1)
            return (redemption + coupon) / (1 + toNext * yield / frequency) - accrued;

        var discount = 1 + yield / frequency;
        var price = redemption / Math.Pow(discount, schedule.Coupons - 1 + toNext);
        for (var k = 1; k <= schedule.Coupons; k++)
            price += coupon / Math.Pow(discount, k - 1 + toNext);
        return price - accrued;
    }

    // YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])
    private static Operand Yield(FunctionCall call)
    {
        if (!TryBond(call, pricing: false, out var settlement, out var maturity, out var rate, out var price, out var redemption,
                out var frequency, out var basis, out var error))
            return error;
        var schedule = Schedule.Of(settlement, maturity, frequency, basis);

        if (schedule.Coupons == 1)
        {
            // Solved in closed form from the one-coupon price.
            var coupon = rate / frequency;
            var paid = price / 100 + schedule.AccruedDays / schedule.PeriodDays * coupon;
            var received = redemption / 100 + coupon;
            return CellValue.Number((received - paid) / paid * frequency * schedule.PeriodDays / schedule.ToNextCoupon);
        }

        var result = SolveYield(y => BondPrice(schedule, rate, y, redemption, frequency), price, rate);
        return result is { } r ? CellValue.Number(r) : CellValue.Error(ErrorKind.Num);
    }

    // DURATION and MDURATION(settlement, maturity, coupon, yld, frequency, [basis]): Macaulay
    // duration in years, the present-value-weighted mean time of the cash flows; MDURATION divides
    // by 1 + yld/frequency.
    private static Operand Duration(FunctionCall call, bool modified)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required, Required, Required, Required, 0))
            return error;
        var system = call.Context.DateSystem;
        var (rate, yield) = (x[2], x[3]);
        if (!TryDate(x[0], system, out var settlement) || !TryDate(x[1], system, out var maturity)
            || !FinancialSecurities.TryFrequency(x[4], out var frequency) || !FinancialSecurities.TryBasis(x[5], out var basis)
            || rate < 0 || yield < 0 || settlement >= maturity)
            return CellValue.Error(ErrorKind.Num);

        var schedule = Schedule.Of(settlement, maturity, frequency, basis);
        var toNext = schedule.ToNextCoupon / schedule.PeriodDays;
        var coupon = 100 * rate / frequency;
        var discount = 1 + yield / frequency;
        double weighted = 0, total = 0;
        for (var k = 1; k <= schedule.Coupons; k++)
        {
            var time = toNext + k - 1;
            var flow = (k == schedule.Coupons ? coupon + 100 : coupon) / Math.Pow(discount, time);
            weighted += time * flow;
            total += flow;
        }

        if (total == 0)
            return CellValue.Error(ErrorKind.Div0);
        var duration = weighted / total / frequency;
        return CellValue.Number(modified ? duration / discount : duration);
    }

    // ODDFPRICE and ODDFYIELD(settlement, maturity, issue, first_coupon, rate, yld or pr, redemption, frequency, [basis])
    private static Operand OddFirst(FunctionCall call, bool price)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required, Required, Required, Required, Required, Required, Required, 0))
            return error;
        var system = call.Context.DateSystem;
        var (rate, value, redemption) = (x[4], x[5], x[6]);
        if (!TryDate(x[0], system, out var settlement) || !TryDate(x[1], system, out var maturity)
            || !TryDate(x[2], system, out var issue) || !TryDate(x[3], system, out var firstCoupon)
            || !FinancialSecurities.TryFrequency(x[7], out var frequency) || !FinancialSecurities.TryBasis(x[8], out var basis)
            || issue >= firstCoupon || firstCoupon > maturity || settlement >= maturity
            || rate < 0 || (price ? value < 0 : value <= 0) || redemption <= 0)
            return CellValue.Error(ErrorKind.Num);

        double PriceAt(double yield) => OddFirstPrice(settlement, maturity, issue, firstCoupon, rate, yield, redemption, frequency, basis);
        if (price)
            return CellValue.Number(PriceAt(value));
        var result = SolveYield(PriceAt, value, rate);
        return result is { } r ? CellValue.Number(r) : CellValue.Error(ErrorKind.Num);
    }

    // Days between two dates by the basis, never negative.
    private static double Days(FinancialDate start, FinancialDate end, int basis) => Math.Max(FinancialDayCount.DaysBetween(start, end, basis), 0);

    private static double OddFirstPrice(FinancialDate settlement, FinancialDate maturity, FinancialDate issue, FinancialDate firstCoupon,
        double rate, double yield, double redemption, int frequency, int basis)
    {
        var months = FinancialDayCount.MonthsPerPeriod(frequency);
        var periodDays = FinancialDayCount.CouponDays(settlement, firstCoupon, frequency, basis);
        var coupon = 100 * rate / frequency;
        var discount = 1 + yield / frequency;
        var firstPeriodDays = Days(issue, firstCoupon, basis);

        if (firstPeriodDays < periodDays)
        {
            // A short first period: its coupon is prorated, the rest is a regular bond.
            var coupons = FinancialDayCount.CouponCount(settlement, maturity, frequency);
            var toFirst = Days(settlement, firstCoupon, basis) / periodDays;
            var accrued = Days(issue, settlement, basis);
            var result = redemption / Math.Pow(discount, coupons - 1 + toFirst)
                         + coupon * firstPeriodDays / periodDays / Math.Pow(discount, toFirst);
            for (var k = 2; k <= coupons; k++)
                result += coupon / Math.Pow(discount, k - 1 + toFirst);
            return result - accrued / periodDays * coupon;
        }

        // A long first period spans several quasi-coupon periods back from the first coupon; each
        // contributes its share of the coupon and of the accrued interest.
        var quasiPeriods = FinancialDayCount.CouponCount(issue, firstCoupon, frequency);
        var late = firstCoupon;
        double couponShare = 0, accruedShare = 0;
        for (var index = quasiPeriods; index >= 1; index--)
        {
            var early = late.AddMonths(-months);
            var length = basis == 1 ? Days(early, late, basis) : periodDays;
            var counted = index > 1 ? length : Days(issue, late, basis);
            var accrued = Days(FinancialDate.Max(issue, early), FinancialDate.Min(settlement, late), basis);
            couponShare += counted / length;
            accruedShare += accrued / length;
            late = early;
        }

        double toNext;
        if (basis is 2 or 3)
        {
            var next = FinancialDayCount.CouponDates(settlement, firstCoupon, frequency).Next;
            toNext = Days(settlement, next, basis);
        }
        else
        {
            var previous = FinancialDayCount.CouponDates(settlement, firstCoupon, frequency).Previous;
            toNext = periodDays - FinancialDayCount.DaysBetween(previous, settlement, basis);
        }

        var periodsToFirst = QuasiCouponsBefore(firstCoupon, settlement, months);
        var regular = FinancialDayCount.CouponCount(firstCoupon, maturity, frequency);
        var y = toNext / periodDays;
        var price = redemption / Math.Pow(discount, y + periodsToFirst + regular)
                    + coupon * couponShare / Math.Pow(discount, periodsToFirst + y);
        for (var k = 1; k <= regular; k++)
            price += coupon / Math.Pow(discount, k + periodsToFirst + y);
        return price - coupon * accruedShare;
    }

    // Whole quasi-coupon periods from settlement up to the anchor date, stepping forward.
    private static double QuasiCouponsBefore(FinancialDate anchor, FinancialDate settlement, int months)
    {
        var monthEnd = anchor.IsLastDayOfMonth;
        if (!monthEnd && anchor.Month != 2 && anchor.Day > 28 && anchor.Day < FinancialDayCount.DaysInMonth(anchor.Year, anchor.Month))
            monthEnd = settlement.IsLastDayOfMonth;

        var start = settlement.AddMonths(0, monthEnd);
        var coupons = settlement < start ? 1.0 : 0.0;
        for (var front = start.AddMonths(months, monthEnd); front < anchor; front = front.AddMonths(months, monthEnd))
            coupons++;
        return coupons;
    }

    // ODDLPRICE and ODDLYIELD(settlement, maturity, last_interest, rate, yld or pr, redemption, frequency, [basis]):
    // the odd last period is split into quasi-coupon periods from the last interest date; both
    // functions are closed forms.
    private static Operand OddLast(FunctionCall call, bool price)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required, Required, Required, Required, Required, Required, 0))
            return error;
        var system = call.Context.DateSystem;
        var (rate, value, redemption) = (x[3], x[4], x[5]);
        if (!TryDate(x[0], system, out var settlement) || !TryDate(x[1], system, out var maturity)
            || !TryDate(x[2], system, out var lastInterest)
            || !FinancialSecurities.TryFrequency(x[6], out var frequency) || !FinancialSecurities.TryBasis(x[7], out var basis)
            || lastInterest >= maturity || settlement >= maturity
            || rate < 0 || (price ? value < 0 : value <= 0) || redemption <= 0)
            return CellValue.Error(ErrorKind.Num);

        var months = FinancialDayCount.MonthsPerPeriod(frequency);
        var periods = FinancialDayCount.CouponCount(lastInterest, maturity, frequency);
        var early = lastInterest;
        double couponShare = 0, accruedShare = 0, remainingShare = 0;
        for (var index = 1; index <= periods; index++)
        {
            var late = early.AddMonths(months);
            // Basis 0 counts these periods by 30/360 with both ends moved.
            var length = basis == 0 ? Math.Max(FinancialDayCount.Days360Us(early, late, modifyBoth: true), 0) : Days(early, late, basis);
            var counted = index < periods ? length
                : basis == 0 ? Math.Max(FinancialDayCount.Days360Us(early, maturity, modifyBoth: true), 0) : Days(early, maturity, basis);
            var accrued = late < settlement ? counted : early < settlement ? Days(early, settlement, basis) : 0;
            var remaining = Days(FinancialDate.Max(settlement, early), FinancialDate.Min(maturity, late), basis);
            couponShare += counted / length;
            accruedShare += accrued / length;
            remainingShare += remaining / length;
            early = late;
        }

        var coupon = 100 * rate / frequency;
        var atMaturity = couponShare * coupon + redemption;
        if (price)
            return CellValue.Number(atMaturity / (remainingShare * value / frequency + 1) - accruedShare * coupon);
        var paid = accruedShare * coupon + value;
        return CellValue.Number((atMaturity - paid) / paid * frequency / remainingShare);
    }

    /// <summary>The yield at which <paramref name="priceAt"/> gives <paramref name="price"/>, by Newton's method from the coupon rate.</summary>
    private static double? SolveYield(Func<double, double> priceAt, double price, double guess)
    {
        const double step = 1e-6;
        var yield = guess;
        for (var i = 0; i < YieldIterations; i++)
        {
            var difference = priceAt(yield) - price;
            var slope = (priceAt(yield + step) - priceAt(yield - step)) / (2 * step);
            var next = yield - difference / slope;
            if (!double.IsFinite(next) || next <= -1)
                return null;
            if (Math.Abs(next - yield) < YieldTolerance)
                return next;
            yield = next;
        }

        return null;
    }
}
