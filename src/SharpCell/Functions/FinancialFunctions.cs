using System;
using SharpCell.Evaluation;
using static SharpCell.Functions.FinancialArguments;

namespace SharpCell.Functions;

/// <summary>
/// Annuities and interest rates. The time-value functions share one equation:
/// <c>pv·(1+r)^n + pmt·(1+r·type)·((1+r)^n − 1)/r + fv = 0</c>, or <c>pv + pmt·n + fv = 0</c> at rate 0.
/// </summary>
internal static class FinancialFunctions
{
    private static readonly ArgumentKind[] Scalars = [ArgumentKind.Value];

    // RATE is solved by Newton's method from the guess, done when a step is below 1e-7. Excel
    // documents 20 steps, but it solves cases that take this method 26 and gives up on ones that
    // take 238; 50 matches every case seen.
    private const int RateIterations = 50;
    private const double RateTolerance = 1e-7;

    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("PV", 3, 5, Scalars, PresentValue));
        registry.Add(new FunctionInfo("FV", 3, 5, Scalars, FutureValue));
        registry.Add(new FunctionInfo("PMT", 3, 5, Scalars, Payment));
        registry.Add(new FunctionInfo("NPER", 3, 5, Scalars, Periods));
        registry.Add(new FunctionInfo("RATE", 3, 6, Scalars, Rate));
        registry.Add(new FunctionInfo("IPMT", 4, 6, Scalars, call => PaymentPart(call, interest: true)));
        registry.Add(new FunctionInfo("PPMT", 4, 6, Scalars, call => PaymentPart(call, interest: false)));
        registry.Add(new FunctionInfo("CUMIPMT", 6, 6, Scalars, call => Cumulative(call, interest: true)));
        registry.Add(new FunctionInfo("CUMPRINC", 6, 6, Scalars, call => Cumulative(call, interest: false)));
        registry.Add(new FunctionInfo("ISPMT", 4, 4, Scalars, InterestOfStraightLoan));
        registry.Add(new FunctionInfo("PDURATION", 3, 3, Scalars, PeriodsToValue));
        registry.Add(new FunctionInfo("RRI", 3, 3, Scalars, EquivalentRate));
        registry.Add(new FunctionInfo("EFFECT", 2, 2, Scalars, Effective));
        registry.Add(new FunctionInfo("NOMINAL", 2, 2, Scalars, Nominal));
        registry.Add(new FunctionInfo("DOLLARDE", 2, 2, Scalars, call => Dollar(call, toDecimal: true)));
        registry.Add(new FunctionInfo("DOLLARFR", 2, 2, Scalars, call => Dollar(call, toDecimal: false)));
    }

    // PV(rate, nper, pmt, [fv], [type])
    private static Operand PresentValue(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required, 0, 0))
            return error;
        var (rate, nper, pmt, fv, start) = (x[0], x[1], x[2], x[3], x[4] != 0);
        if (rate == 0)
            return CellValue.Number(-fv - pmt * nper);
        if (rate == -1)
            return CellValue.Error(ErrorKind.Div0);

        var growth = Math.Pow(1 + rate, nper);
        var result = start
            ? -(fv * rate + pmt * (1 + rate) * (growth - 1)) / (rate * growth)
            : (-fv * rate - pmt * (growth - 1)) / (rate * growth);
        return CellValue.Number(result);
    }

    // FV(rate, nper, pmt, [pv], [type])
    private static Operand FutureValue(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required, 0, 0))
            return error;
        var (rate, nper) = (x[0], x[1]);
        if (rate == -1 && nper < 0)
            return CellValue.Error(ErrorKind.Div0);
        var result = Future(rate, nper, x[2], x[3], x[4] != 0);
        return double.IsInfinity(result) ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(result);
    }

    private static double Future(double rate, double nper, double pmt, double pv, bool start)
    {
        if (rate == 0)
            return -pv - pmt * nper;
        var growth = Math.Pow(1 + rate, nper);
        return start
            ? -pv * growth - pmt * (1 + rate) * (growth - 1) / rate
            : -pv * growth - pmt * (growth - 1) / rate;
    }

    // PMT(rate, nper, pv, [fv], [type])
    private static Operand Payment(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required, 0, 0))
            return error;
        return CellValue.Number(Payment(x[0], x[1], x[2], x[3], x[4] != 0));
    }

    // NaN where Excel gives #NUM!.
    private static double Payment(double rate, double nper, double pv, double fv, bool start)
    {
        if (rate == 0)
            return nper == 0 ? double.NaN : -(pv + fv) / nper;
        if (rate <= -1)
            return double.NaN;

        var growth = nper == 0 ? 1 : Math.Pow(1 + rate, nper);
        return start
            ? (fv + pv * growth) * rate / ((1 + rate) * (1 - growth))
            : (fv * rate + pv * rate * growth) / (1 - growth);
    }

    // NPER(rate, pmt, pv, [fv], [type])
    private static Operand Periods(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required, 0, 0))
            return error;
        var (rate, pmt, pv, fv, start) = (x[0], x[1], x[2], x[3], x[4] != 0);
        if (rate == 0)
            return pmt == 0 ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(-(fv + pv) / pmt);
        if (rate < -1)
            return CellValue.Error(ErrorKind.Num);

        // (1+r)^n from the equation, then its logarithm.
        double growth;
        if (pmt == 0)
        {
            growth = -fv / pv;
        }
        else
        {
            var term = start ? pmt * (1 + rate) / rate : pmt / rate;
            growth = (1 - fv / term) / (1 + pv / term);
        }

        if (!(growth > 0))
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(Math.Log(growth) / Math.Log(1 + rate));
    }

    // RATE(nper, pmt, pv, [fv], [type], [guess])
    private static Operand Rate(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required, 0, 0, 0.1))
            return error;
        var (nper, pmt, pv, fv, type, guess) = (x[0], x[1], x[2], x[3], x[4] != 0 ? 1.0 : 0.0, x[5]);
        if (guess <= -1)
            return CellValue.Error(ErrorKind.Value);

        var rate = guess;
        for (var i = 0; i < RateIterations; i++)
        {
            double f, slope;
            if (rate == 0)
            {
                f = pv + pmt * nper + fv;
                slope = pv * nper + pmt * (nper * (nper - 1) / 2 + type * nper);
            }
            else
            {
                var before = Math.Pow(1 + rate, nper - 1);
                var growth = before * (1 + rate);
                f = pv * growth + pmt * (1 + rate * type) * (growth - 1) / rate + fv;
                slope = pv * nper * before
                        + pmt * type * (growth - 1) / rate
                        + pmt * (1 + rate * type) * (nper * before * rate - (growth - 1)) / (rate * rate);
            }

            var next = rate - f / slope;
            if (!double.IsFinite(next) || next <= -1)
                return CellValue.Error(ErrorKind.Num);
            if (Math.Abs(next - rate) < RateTolerance)
                return CellValue.Number(next);
            rate = next;
        }

        return CellValue.Error(ErrorKind.Num);
    }

    // IPMT and PPMT(rate, per, nper, pv, [fv], [type]): the interest and principal parts of one payment.
    private static Operand PaymentPart(FunctionCall call, bool interest)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required, Required, 0, 0))
            return error;
        var (rate, period, nper, pv, fv, start) = (x[0], x[1], x[2], x[3], x[4], x[5] != 0);
        var payment = Payment(rate, nper, pv, fv, start);
        if (double.IsNaN(payment))
            return CellValue.Error(ErrorKind.Num);
        if (period < 1 || period >= nper + 1)
            return CellValue.Error(ErrorKind.Num);

        var part = InterestPart(rate, period, payment, pv, start);
        if (double.IsInfinity(part))
            return CellValue.Error(ErrorKind.Div0);
        return CellValue.Number(interest ? part : payment - part);
    }

    // The interest in the payment of a period: the rate on the balance after the previous period.
    private static double InterestPart(double rate, double period, double payment, double pv, bool start)
    {
        if (start && period == 1)
            return 0;
        return start
            ? (Future(rate, period - 2, payment, pv, start) - payment) * rate
            : Future(rate, period - 1, payment, pv, start) * rate;
    }

    // CUMIPMT and CUMPRINC(rate, nper, pv, start_period, end_period, type)
    private static Operand Cumulative(FunctionCall call, bool interest)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required, Required, Required, Required, Required))
            return error;
        var (rate, nper, pv, type) = (x[0], x[1], x[2], x[5]);
        var first = Math.Ceiling(x[3]);
        var last = Math.Truncate(x[4]);
        if (type is not (0 or 1) || first > last || rate <= 0 || nper <= 0 || pv <= 0 || first < 1)
            return CellValue.Error(ErrorKind.Num);

        var start = type == 1;
        var payment = Payment(rate, nper, pv, 0, start);
        if (double.IsNaN(payment) || last >= nper + 1)
            return CellValue.Error(ErrorKind.Num);

        // The principal of period k is r·pv/(1 − (1+r)^nper)·(1+r)^(k−1), a period earlier when paid
        // in advance, where the first payment is all principal. Summing these, rather than payment
        // minus interest, keeps the tiny principals of a high rate exact, as Excel does.
        var cancellation = call.Context.CancellationToken;
        var principalBase = rate * pv / (1 - Math.Pow(1 + rate, nper));
        var principal = 0.0;
        for (var period = first; period <= last; period++)
        {
            if ((long)(period - first) % 4096 == 4095)
                cancellation.ThrowIfCancellationRequested();
            principal += start && period == 1 ? payment : principalBase * Math.Pow(1 + rate, start ? period - 2 : period - 1);
        }

        return CellValue.Number(interest ? (last - first + 1) * payment - principal : principal);
    }

    // ISPMT(rate, per, nper, pv): the interest of a period of a loan repaid in equal principal parts.
    private static Operand InterestOfStraightLoan(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required, Required))
            return error;
        var (rate, period, nper, pv) = (x[0], x[1], x[2], x[3]);
        if (nper == 0)
            return CellValue.Error(ErrorKind.Div0);
        return CellValue.Number(pv * rate * (period / nper - 1));
    }

    // PDURATION(rate, pv, fv): periods for pv to grow to fv.
    private static Operand PeriodsToValue(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required))
            return error;
        var (rate, pv, fv) = (x[0], x[1], x[2]);
        if (rate <= 0 || pv <= 0 || fv <= 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number((Math.Log(fv) - Math.Log(pv)) / Math.Log(1 + rate));
    }

    // RRI(nper, pv, fv): the rate per period that grows pv to fv.
    private static Operand EquivalentRate(FunctionCall call)
    {
        if (!TryNumbers(call, out var x, out var error, Required, Required, Required))
            return error;
        var (nper, pv, fv) = (x[0], x[1], x[2]);
        if (nper <= 0)
            return CellValue.Error(ErrorKind.Num);
        if (fv == pv)
            return CellValue.Number(0);
        if (pv == 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(Math.Pow(fv / pv, 1 / nper) - 1);
    }

    // EFFECT(nominal_rate, npery)
    private static Operand Effective(FunctionCall call)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required))
            return error;
        var (nominal, periods) = (x[0], Math.Truncate(x[1]));
        if (nominal <= 0 || periods < 1)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(Math.Pow(1 + nominal / periods, periods) - 1);
    }

    // NOMINAL(effect_rate, npery)
    private static Operand Nominal(FunctionCall call)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required))
            return error;
        var (effective, periods) = (x[0], Math.Truncate(x[1]));
        if (effective <= 0 || periods < 1)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number((Math.Pow(1 + effective, 1 / periods) - 1) * periods);
    }

    // DOLLARDE(fractional_dollar, fraction) and DOLLARFR(decimal_dollar, fraction): 1.02 in
    // sixteenths is 1 2/16 = 1.125. The fraction's digits take as many decimal places as the
    // denominator has digits.
    private static Operand Dollar(FunctionCall call, bool toDecimal)
    {
        if (!TryAddInNumbers(call, out var x, out var error, Required, Required))
            return error;
        var (dollar, fraction) = (x[0], x[1]);
        if (fraction < 0)
            return CellValue.Error(ErrorKind.Num);
        var denominator = Math.Truncate(fraction);
        if (denominator < 1)
            return CellValue.Error(ErrorKind.Div0);

        var whole = Math.Truncate(dollar);
        var rest = dollar - whole;
        var scale = Math.Pow(10, Math.Ceiling(Math.Log10(denominator)));
        return CellValue.Number(toDecimal ? rest * scale / denominator + whole : rest * denominator / scale + whole);
    }
}
