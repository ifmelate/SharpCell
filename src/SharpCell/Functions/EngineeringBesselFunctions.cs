using System;
using System.Threading;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// BESSELI, BESSELJ, BESSELK and BESSELY for integer orders (Excel truncates the order). The
/// methods aim at full double precision, where Excel's own approximations are good to about
/// eight digits:
/// <list type="bullet">
/// <item>J by Miller's backward recurrence, normalized by J0 + 2(J2 + J4 + ...) = 1;</item>
/// <item>Y from the same recurrence through Neumann's series for Y0 and Y1, then forward recurrence;</item>
/// <item>J and Y for a large argument and a small order by Hankel's asymptotic expansion;</item>
/// <item>I by its power series, which has only positive terms;</item>
/// <item>K by the trapezoidal rule on K_n(x) = ∫ exp(-x cosh t) cosh(nt) dt over t ≥ 0, which converges
/// exponentially fast for this integrand.</item>
/// </list>
/// </summary>
internal static class EngineeringBesselFunctions
{
    private const double EulerGamma = 0.57721566490153286;

    // No method runs more steps than this; beyond it (orders or arguments in the tens of millions
    // that Hankel's expansion does not cover) the result is #NUM!.
    private const int MaxSteps = 50_000_000;

    private const int CancellationCheckInterval = 1 << 16;

    // Hankel's expansion is used from this argument on, for orders with n² below x / 100.
    private const double AsymptoticArgument = 1e5;

    // Values are rescaled when they pass this size during a recurrence or a sum.
    private const double Rescale = 1e250;

    private static readonly ArgumentKind[] Values = [ArgumentKind.Value];

    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("BESSELI", 2, 2, Values, call => Evaluate(call, I, secondKind: false)));
        registry.Add(new FunctionInfo("BESSELJ", 2, 2, Values, call => Evaluate(call, J, secondKind: false))
        {
            Status = FunctionStatus.KnownDeviation,
            Deviation = "Accurate to full double precision, while Excel's approximation is off by up to about 5E-7 (BESSELJ(10,1) is 0.0434722 in Excel, 0.0434727 exactly).",
        });
        registry.Add(new FunctionInfo("BESSELK", 2, 2, Values, call => Evaluate(call, K, secondKind: true)));
        registry.Add(new FunctionInfo("BESSELY", 2, 2, Values, call => Evaluate(call, Y, secondKind: true)));
    }

    // A negative order is #NUM!, and so is x ≤ 0 for the functions of the second kind, K and Y. A
    // result that is not finite (overflow, or more steps than allowed, reported as NaN) is #NUM!.
    private static Operand Evaluate(FunctionCall call, Func<double, int, CancellationToken, double> bessel, bool secondKind)
    {
        var x = EngineeringFunctions.StrictNumber(call, 0);
        if (x.IsError)
            return x;
        var order = EngineeringFunctions.StrictNumber(call, 1);
        if (order.IsError)
            return order;

        var n = Math.Truncate(order.AsNumber());
        var value = x.AsNumber();
        if (n < 0 || (secondKind && value <= 0))
            return CellValue.Error(ErrorKind.Num);

        // I and J vanish at high orders, long before the step limit: |J_n(x)| ≤ (x/2)^n / n! and
        // I_n(x) ≤ (x/2)^n / n! · exp(x² / (4(n + 1))).
        if (!secondKind && value != 0 && n * Math.Log(Math.Abs(value) / 2) - LogFactorial(n) + value * value / (4 * (n + 1)) < -760)
            return CellValue.Number(0);
        if (n > MaxSteps)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(bessel(value, (int)n, call.Context.CancellationToken));
    }

    private static double J(double x, int n, CancellationToken cancellation)
    {
        if (x == 0)
            return n == 0 ? 1 : 0;

        // J_n(-x) = (-1)^n J_n(x).
        var sign = x < 0 && n % 2 == 1 ? -1 : 1;
        x = Math.Abs(x);
        if (UsesHankel(x, n))
            return sign * Hankel(x, n).J;
        return sign * Miller(x, n, cancellation).Jn;
    }

    private static double Y(double x, int n, CancellationToken cancellation)
    {
        if (UsesHankel(x, n))
            return Hankel(x, n).Y;

        // Only J0, J1 and the Neumann sums are needed, so the recurrence need not start above n.
        var m = Miller(x, 1, cancellation);
        var logTerm = Math.Log(x / 2) + EulerGamma;
        var y0 = 2 / Math.PI * logTerm * m.J0 - 4 / Math.PI * m.Neumann0;
        if (n == 0)
            return y0;

        // ψ(2) = 1 - γ, so ln(x/2) - ψ(2) = ln(x/2) + γ - 1.
        var y1 = -2 / (Math.PI * x) * m.J0 + 2 / Math.PI * (logTerm - 1) * m.J1 - 2 / Math.PI * m.Neumann1;
        return ForwardY(x, n, y0, y1, cancellation);
    }

    private static double ForwardY(double x, int n, double y0, double y1, CancellationToken cancellation)
    {
        double previous = y0, current = y1;
        for (var k = 1; k < n && double.IsFinite(current); k++)
        {
            if (k % CancellationCheckInterval == 0)
                cancellation.ThrowIfCancellationRequested();
            (previous, current) = (current, 2.0 * k / x * current - previous);
        }

        return n == 0 ? y0 : current;
    }

    private static bool UsesHankel(double x, int n) => x >= AsymptoticArgument && (double)n * n < x / 100;

    // J_n and Y_n ≈ sqrt(2/(πx)) (P cos χ ∓ Q sin χ), χ = x - (n/2 + 1/4)π, with P and Q the
    // asymptotic series in 1/(8x). For n² < x/100 the terms fall below double precision quickly.
    private static (double J, double Y) Hankel(double x, int n)
    {
        var mu = 4.0 * n * n;
        double p = 0, q = 0, term = 1;
        for (var k = 0; k < 60 && Math.Abs(term) > 1e-17; k++)
        {
            if (k % 4 == 0)
                p += term;
            else if (k % 4 == 1)
                q += term;
            else if (k % 4 == 2)
                p -= term;
            else
                q -= term;
            var odd = 2 * k + 1;
            term *= (mu - (double)odd * odd) / ((k + 1) * 8 * x);
        }

        // cos(x - φ) and sin(x - φ) from the exact sine and cosine of x: subtracting φ from a large
        // x first would round away the phase.
        var phi = (n / 2.0 + 0.25) * Math.PI;
        var (sinX, cosX) = Math.SinCos(x);
        var cosChi = cosX * Math.Cos(phi) + sinX * Math.Sin(phi);
        var sinChi = sinX * Math.Cos(phi) - cosX * Math.Sin(phi);
        var scale = Math.Sqrt(2 / (Math.PI * x));
        return (scale * (p * cosChi - q * sinChi), scale * (p * sinChi + q * cosChi));
    }

    private readonly record struct MillerResult(double Jn, double J0, double J1, double Neumann0, double Neumann1);

    /// <summary>
    /// Miller's algorithm for x &gt; 0: recurse J_{k-1} = (2k/x) J_k - J_{k+1} down from an index far
    /// above both n and x, then normalize. Along the way it collects the sums of Neumann's series,
    /// Σ (-1)^k J_2k / k and Σ (-1)^k (2k+1) J_(2k+1) / (k(k+1)) for k ≥ 1, which give Y0 and Y1.
    /// </summary>
    private static MillerResult Miller(double x, int n, CancellationToken cancellation)
    {
        var big = Math.Max(n, x);
        var start = big + 20 + Math.Sqrt(160 * big);
        if (start > MaxSteps)
            return new MillerResult(double.NaN, double.NaN, double.NaN, double.NaN, double.NaN);

        var top = 2 * ((int)start / 2 + 1);
        double next = 0, current = 1;
        double jn = 0, j1 = 0, norm = 0, neumann0 = 0, neumann1 = 0;
        Collect(top, current);
        for (var k = top; k >= 1; k--)
        {
            if (k % CancellationCheckInterval == 0)
                cancellation.ThrowIfCancellationRequested();
            (next, current) = (current, 2.0 * k / x * current - next);
            Collect(k - 1, current);
            if (Math.Abs(current) > Rescale)
            {
                const double factor = 1 / Rescale;
                current *= factor;
                next *= factor;
                jn *= factor;
                j1 *= factor;
                norm *= factor;
                neumann0 *= factor;
                neumann1 *= factor;
            }
        }

        return new MillerResult(jn / norm, current / norm, j1 / norm, neumann0 / norm, neumann1 / norm);

        void Collect(int index, double value)
        {
            if (index == n)
                jn = value;
            if (index == 1)
                j1 = value;
            if (index == 0)
            {
                norm += value;
            }
            else if (index % 2 == 0)
            {
                var half = index / 2;
                norm += 2 * value;
                neumann0 += (half % 2 == 0 ? value : -value) / half;
            }
            else if (index >= 3)
            {
                var half = (index - 1) / 2;
                var term = (double)index * value / ((double)half * (half + 1));
                neumann1 += half % 2 == 0 ? term : -term;
            }
        }
    }

    // I_n(x) = Σ (x/2)^(2k+n) / (k! (n+k)!), summed relative to its first term with the scale kept
    // as a logarithm, so that neither the first term nor the sum overflows on the way.
    private static double I(double x, int n, CancellationToken cancellation)
    {
        if (x == 0)
            return n == 0 ? 1 : 0;

        var sign = x < 0 && n % 2 == 1 ? -1 : 1;
        x = Math.Abs(x);
        var quarter = x * x / 4;
        double term = 1, sum = 1, logScale = 0;
        for (var k = 0; ; k++)
        {
            if (k >= MaxSteps)
                return double.NaN;
            if (k % CancellationCheckInterval == 0 && k > 0)
                cancellation.ThrowIfCancellationRequested();
            term *= quarter / ((k + 1.0) * (n + k + 1.0));
            sum += term;
            if (term < sum * 1e-17 && k + 1 > Math.Sqrt(quarter))
                break;
            if (sum > Rescale)
            {
                logScale += Math.Log(sum);
                term /= sum;
                sum = 1;
            }
        }

        return sign * Math.Exp(n * Math.Log(x / 2) - LogFactorial(n) + logScale + Math.Log(sum));
    }

    // K_n(x) = ∫ exp(-x cosh t) cosh(nt) dt over t ≥ 0. The integrand is smooth and decays doubly
    // exponentially, so the trapezoidal rule converges exponentially; the step follows the width
    // of the peak, about (x² + n²)^(-1/4). Terms are summed as logarithms against the largest one.
    private static double K(double x, int n, CancellationToken cancellation)
    {
        var step = Math.Min(0.1, 0.5 / Math.Pow(x * x + (double)n * n, 0.25));
        double max = double.NegativeInfinity, sum = 0;
        for (var i = 0; ; i++)
        {
            if (i >= MaxSteps)
                return double.NaN;
            if (i % CancellationCheckInterval == 0 && i > 0)
                cancellation.ThrowIfCancellationRequested();

            var t = i * step;
            var log = -x * Math.Cosh(t) + LogCosh(n * t);
            var weight = i == 0 ? 0.5 : 1;
            if (log > max)
            {
                sum = sum * Math.Exp(max - log) + weight;
                max = log;
            }
            else
            {
                sum += weight * Math.Exp(log - max);
            }

            // Past the peak (where x sinh t = n tanh nt ≤ n) and negligible.
            if (log < max - 45 && x * Math.Sinh(t) > n)
                break;
        }

        return Math.Exp(max + Math.Log(sum * step));
    }

    private static double LogCosh(double u) => u < 20 ? Math.Log(Math.Cosh(u)) : u - Math.Log(2) + Math.Log(1 + Math.Exp(-2 * u));

    private static double LogFactorial(double n)
    {
        if (n < 2)
            return 0;
        if (n <= 170)
        {
            var product = 1.0;
            for (var k = 2.0; k <= n; k++)
                product *= k;
            return Math.Log(product);
        }

        // Stirling's series; at n > 170 the first terms give full double precision.
        return n * Math.Log(n) - n + 0.5 * Math.Log(2 * Math.PI * n) + 1 / (12 * n) - 1 / (360 * n * n * n);
    }
}
