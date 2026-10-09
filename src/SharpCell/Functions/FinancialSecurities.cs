using System;
using SharpCell.Evaluation;
using static SharpCell.Functions.FinancialArguments;

namespace SharpCell.Functions;

/// <summary>
/// Securities without a coupon schedule to price: accrued interest (ACCRINT, ACCRINTM), discounted
/// securities (DISC, INTRATE, RECEIVED, PRICEDISC, YIELDDISC), securities paying interest at maturity
/// (PRICEMAT, YIELDMAT) and Treasury bills. Dates are truncated to whole days; a basis or frequency
/// out of range, or dates in the wrong order, give <c>#NUM!</c>.
/// </summary>
internal static class FinancialSecurities
{
    private static readonly ArgumentKind[] Scalars = [ArgumentKind.Value];

    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("ACCRINT", 6, 8, Scalars, AccruedInterest));
        registry.Add(new FunctionInfo("ACCRINTM", 3, 5, Scalars, AccruedInterestAtMaturity));
        registry.Add(new FunctionInfo("DISC", 4, 5, Scalars, call => Discounted(call, Discount)));
        registry.Add(new FunctionInfo("INTRATE", 4, 5, Scalars, call => Discounted(call, InterestRate)));
        registry.Add(new FunctionInfo("RECEIVED", 4, 5, Scalars, call => Discounted(call, Received)));
        registry.Add(new FunctionInfo("PRICEDISC", 4, 5, Scalars, call => Discounted(call, PriceOfDiscounted)));
        registry.Add(new FunctionInfo("YIELDDISC", 4, 5, Scalars, call => Discounted(call, YieldOfDiscounted)));
        registry.Add(new FunctionInfo("PRICEMAT", 5, 6, Scalars, call => AtMaturity(call, price: true)));
        registry.Add(new FunctionInfo("YIELDMAT", 5, 6, Scalars, call => AtMaturity(call, price: false)));
        registry.Add(new FunctionInfo("TBILLEQ", 3, 3, Scalars, call => TreasuryBill(call, BondEquivalentYield)));
        registry.Add(new FunctionInfo("TBILLPRICE", 3, 3, Scalars, call => TreasuryBill(call, BillPrice)));
        registry.Add(new FunctionInfo("TBILLYIELD", 3, 3, Scalars, call => TreasuryBill(call, BillYield)));
    }

    /// <summary>A basis argument: truncated, 0 to 4.</summary>
    public static bool TryBasis(double value, out int basis)
    {
        var whole = Math.Truncate(value);
        basis = whole is >= 0 and <= 4 ? (int)whole : -1;
        return basis >= 0;
    }

    /// <summary>A frequency argument: truncated, 1, 2 or 4 coupons a year.</summary>
    public static bool TryFrequency(double value, out int frequency)
    {
        var whole = Math.Truncate(value);
        frequency = whole is 1 or 2 or 4 ? (int)whole : 0;
        return frequency != 0;
    }

    // ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])
    private static Operand AccruedInterest(FunctionCall call)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required, Required, Required, 1000, Required, 0))
            return error;
        var calcMethod = call.Boolean(7, true);
        if (calcMethod.IsError)
            return calcMethod;

        var system = call.Context.DateSystem;
        var (rate, par) = (x[3], x[4]);
        if (!TryDate(x[0], system, out var issue) || !TryDate(x[1], system, out var firstInterest)
            || !TryDate(x[2], system, out var settlement) || !TryFrequency(x[5], out var frequency) || !TryBasis(x[6], out var basis)
            || rate <= 0 || par <= 0 || issue >= settlement)
            return CellValue.Error(ErrorKind.Num);

        return CellValue.Number(par * rate / frequency
                                * AccruedPeriods(issue, firstInterest, settlement, frequency, basis, calcMethod.AsBoolean()));
    }

    // The accrual in coupon periods. The quasi-coupon periods run from first_interest in steps of
    // 12/frequency months (on month ends when first_interest is one). The period holding settlement
    // counts its prorated part; earlier whole periods count 1 when accruing from issue (calc_method
    // TRUE) and 0 when from first_interest; the period holding issue counts its prorated part.
    private static double AccruedPeriods(FinancialDate issue, FinancialDate firstInterest, FinancialDate settlement, int frequency,
        int basis, bool fromIssue)
    {
        var months = FinancialDayCount.MonthsPerPeriod(frequency);
        var monthEnd = firstInterest.IsLastDayOfMonth;
        var regularStart = firstInterest.AddMonths(-months, monthEnd);

        // The last coupon date before settlement.
        var previous = regularStart;
        if (settlement > firstInterest && fromIssue)
        {
            for (var date = firstInterest; date < settlement; date = date.AddMonths(months, monthEnd))
                previous = date;
        }

        var start = FinancialDate.Max(issue, previous);
        var accrued = FinancialDayCount.DaysBetween(start, settlement, basis)
                      / FinancialDayCount.CouponDays(regularStart, firstInterest, frequency, basis);

        var end = previous;
        while (end > issue)
        {
            var begin = end.AddMonths(-months, monthEnd);
            if (issue <= begin)
            {
                accrued += fromIssue ? 1 : 0;
            }
            else
            {
                var days = basis == 0 ? FinancialDayCount.Days360Us(issue, end) : FinancialDayCount.DaysBetween(issue, end, basis);
                var length = basis switch
                {
                    0 => FinancialDayCount.Days360Us(begin, end, modifyBoth: true),
                    3 => 365.0 / frequency,
                    _ => FinancialDayCount.DaysBetween(begin, end, basis, numerator: false),
                };
                accrued += days / length;
            }

            end = begin;
        }

        return accrued;
    }

    // ACCRINTM(issue, settlement, rate, [par], [basis]): interest paid at maturity, accrued to settlement.
    private static Operand AccruedInterestAtMaturity(FunctionCall call)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required, Required, 1000, 0))
            return error;
        var system = call.Context.DateSystem;
        var (rate, par) = (x[2], x[3]);
        if (!TryDate(x[0], system, out var issue) || !TryDate(x[1], system, out var settlement) || !TryBasis(x[4], out var basis)
            || rate <= 0 || par <= 0 || issue >= settlement)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(par * rate * FinancialDayCount.YearFraction(issue, settlement, basis));
    }

    // The discounted-security functions: (settlement, maturity, a, b, [basis]) with a and b positive,
    // given the year fraction from settlement to maturity.
    private static Operand Discounted(FunctionCall call, Func<double, double, double, double> formula)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required, Required, Required, 0))
            return error;
        var system = call.Context.DateSystem;
        var (a, b) = (x[2], x[3]);
        if (!TryDate(x[0], system, out var settlement) || !TryDate(x[1], system, out var maturity) || !TryBasis(x[4], out var basis)
            || a <= 0 || b <= 0 || settlement >= maturity)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(formula(a, b, FinancialDayCount.YearFraction(settlement, maturity, basis)));
    }

    // DISC(settlement, maturity, pr, redemption, [basis])
    private static double Discount(double price, double redemption, double years) => (1 - price / redemption) / years;

    // INTRATE(settlement, maturity, investment, redemption, [basis])
    private static double InterestRate(double investment, double redemption, double years) => (redemption / investment - 1) / years;

    // RECEIVED(settlement, maturity, investment, discount, [basis]); NaN (#NUM!) when the discount eats it all.
    private static double Received(double investment, double discount, double years)
    {
        var share = 1 - discount * years;
        return share <= 0 ? double.NaN : investment / share;
    }

    // PRICEDISC(settlement, maturity, discount, redemption, [basis])
    private static double PriceOfDiscounted(double discount, double redemption, double years) => redemption * (1 - discount * years);

    // YIELDDISC(settlement, maturity, pr, redemption, [basis])
    private static double YieldOfDiscounted(double price, double redemption, double years) => (redemption / price - 1) / years;

    // PRICEMAT and YIELDMAT(settlement, maturity, issue, rate, yld or pr, [basis]): the three spans
    // from issue, settlement and maturity share the year length of the issue-to-settlement span.
    private static Operand AtMaturity(FunctionCall call, bool price)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required, Required, Required, Required, 0))
            return error;
        var system = call.Context.DateSystem;
        var (rate, value) = (x[3], x[4]);
        if (!TryDate(x[0], system, out var settlement) || !TryDate(x[1], system, out var maturity) || !TryDate(x[2], system, out var issue)
            || !TryBasis(x[5], out var basis) || rate < 0 || (price ? value < 0 : value <= 0)
            || settlement >= maturity || issue >= settlement)
            return CellValue.Error(ErrorKind.Num);

        var year = FinancialDayCount.DaysInYear(issue, settlement, basis);
        var issueToSettlement = FinancialDayCount.DaysBetween(issue, settlement, basis) / year;
        var issueToMaturity = FinancialDayCount.DaysBetween(issue, maturity, basis) / year;
        var settlementToMaturity = issueToMaturity - issueToSettlement;

        if (price)
        {
            var yield = value;
            return CellValue.Number(100 * ((1 + rate * issueToMaturity) / (1 + yield * settlementToMaturity) - rate * issueToSettlement));
        }

        var dirty = value / 100 + rate * issueToSettlement;
        if (dirty == 0 || settlementToMaturity == 0)
            return CellValue.Error(ErrorKind.Div0);
        return CellValue.Number(((1 + rate * issueToMaturity) / dirty - 1) / settlementToMaturity);
    }

    // TBILLEQ, TBILLPRICE and TBILLYIELD(settlement, maturity, x): maturity no later than a year after
    // settlement, x positive; days counted actual/360.
    private static Operand TreasuryBill(FunctionCall call, Func<double, double, double> formula)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required, Required))
            return error;
        var system = call.Context.DateSystem;
        if (!TryDate(x[0], system, out var settlement) || !TryDate(x[1], system, out var maturity)
            || settlement > maturity || maturity > settlement.AddMonths(12) || x[2] <= 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(formula(FinancialDayCount.Actual(settlement, maturity), x[2]));
    }

    // The bond-equivalent yield. Under half a year it is simple interest on the price; beyond, the
    // bill is compared with a semiannual bond, which gives a quadratic in the yield.
    private static double BondEquivalentYield(double days, double discount)
    {
        if (days < 183)
            return 365 * discount / (360 - discount * days);

        var year = days == 366 ? 366.0 : 365.0;
        var extra = days - year / 2;
        var price = 1 - days * discount / 360;
        var a = extra * price / (year * 2);
        var b = price * (0.5 + extra / year);
        var c = price - 1;
        return (-b + Math.Sqrt(b * b - 4 * a * c)) / (2 * a);
    }

    private static double BillPrice(double days, double discount)
    {
        var price = 100 * (1 - discount * days / 360);
        return price < 0 ? double.NaN : price;
    }

    private static double BillYield(double days, double price) => (100 - price) * 360 / (price * days);
}
