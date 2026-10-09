using System;

namespace SharpCell.Functions;

/// <summary>
/// The special functions behind the statistical distributions: the error function, log-gamma and
/// gamma, regularized incomplete gamma and beta functions with their inverses, and the normal
/// distribution. Accurate to a few units in the last place where the problem is well conditioned.
/// Nothing here throws: results outside the double range come back as NaN or infinity, which
/// callers turn into <c>#NUM!</c>. Every iteration is capped.
/// </summary>
internal static class Special
{
    private const double Epsilon = 2.220446049250313e-16;
    private const double Tiny = 1e-300;
    private const double SqrtPi = 1.7724538509055160273;
    private const double SqrtTwoPi = 2.5066282746310005024;
    private const double LogTwoPi = 1.8378770664093454836;
    private const double LogSqrtTwoPi = 0.91893853320467274178;
    private const double EulerGamma = 0.57721566490153286061;

    // Enough for the series and continued fractions to converge for parameters up to about 1e11:
    // near the mean they need a few times the square root of the shape parameter.
    private const int MaxIterations = 5_000_000;

    // zeta(k) - 1 for k = 2..ZetaTerms-1, the coefficients of the log-gamma series around 1.
    private const int ZetaTerms = 64;
    private static readonly double[] ZetaMinusOne = BuildZetaMinusOne();

    // ---------------------------------------------------------------------------------------
    // Elementary helpers

    /// <summary>log(1 + x), accurate for small x.</summary>
    public static double Log1p(double x)
    {
        var u = 1 + x;
        if (u == 1)
            return x;
        // The rounding error of 1 + x cancels in the ratio (Goldberg).
        return Math.Log(u) * x / (u - 1);
    }

    /// <summary>exp(x) - 1, accurate for small x.</summary>
    public static double Expm1(double x)
    {
        if (Math.Abs(x) < 1e-5)
            return x + x * x / 2 + x * x * x / 6;
        var u = Math.Exp(x);
        if (u == 1)
            return x;
        var um1 = u - 1;
        if (um1 == -1)
            return -1;
        // The rounding error of exp(x) cancels in the ratio (Kahan).
        return double.IsInfinity(u) ? u : um1 * x / Math.Log(u);
    }

    /// <summary>log(1 + x) - x without cancellation for small x.</summary>
    public static double Log1pmx(double x)
    {
        if (Math.Abs(x) >= 0.5)
            return Log1p(x) - x;

        // log(1+x) = 2 atanh(r) with r = x / (2 + x), and x - 2r = r x.
        var r = x / (2 + x);
        var r2 = r * r;
        var power = r * r2;
        var sum = 0.0;
        for (var k = 3; k < 200; k += 2)
        {
            var term = power / k;
            sum += term;
            if (Math.Abs(term) <= Math.Abs(sum) * Epsilon)
                break;
            power *= r2;
        }

        return 2 * sum - r * x;
    }

    /// <summary>sin(pi x) with the argument reduced exactly, so it is zero at the integers.</summary>
    public static double SinPi(double x)
    {
        if (x < 0)
            return -SinPi(-x);
        var r = x % 2;
        if (r == 0 || r == 1)
            return 0;
        if (r > 1)
            return -SinPi(r - 1);
        return r <= 0.5 ? Math.Sin(Math.PI * r) : Math.Sin(Math.PI * (1 - r));
    }

    // ---------------------------------------------------------------------------------------
    // Error function and the normal distribution

    public static double Erf(double x)
    {
        if (double.IsNaN(x))
            return x;
        var a = Math.Abs(x);
        if (a < 1)
            return ErfSeries(x);
        var r = 1 - ErfcLarge(a);
        return x < 0 ? -r : r;
    }

    public static double Erfc(double x)
    {
        if (double.IsNaN(x))
            return x;
        if (x >= 1)
            return ErfcLarge(x);
        if (x > -1)
            return 1 - ErfSeries(x);
        return 2 - ErfcLarge(-x);
    }

    // erf(x) = 2/sqrt(pi) exp(-x^2) sum 2^n x^(2n+1) / (1*3*...*(2n+1)): positive terms only.
    private static double ErfSeries(double x)
    {
        var term = x;
        var sum = x;
        var twoX2 = 2 * x * x;
        for (var n = 1; n < 200; n++)
        {
            term *= twoX2 / (2 * n + 1);
            sum += term;
            if (Math.Abs(term) <= Math.Abs(sum) * 1e-17)
                break;
        }

        return 2 / SqrtPi * ExpMinusSquare(x) * sum;
    }

    // erfc(x) for x >= 1 by its continued fraction.
    private static double ErfcLarge(double x)
    {
        if (x > 27.3)
            return 0;
        return ExpMinusSquare(x) / SqrtPi * 2 * x / TailFraction(2 * x * x);
    }

    // s + 1 - 1*2/(s + 5 - 3*4/(s + 9 - ...)) by the modified Lentz method; erfc(x) with s = 2x^2.
    private static double TailFraction(double s)
    {
        var f = s + 1;
        var c = f;
        var d = 0.0;
        for (var n = 1; n < 10_000; n++)
        {
            var a = -(2.0 * n - 1) * (2.0 * n);
            var b = s + 1 + 4.0 * n;
            d = b + a * d;
            if (d == 0)
                d = Tiny;
            c = b + a / c;
            if (c == 0)
                c = Tiny;
            d = 1 / d;
            var delta = c * d;
            f *= delta;
            if (Math.Abs(delta - 1) < Epsilon)
                break;
        }

        return f;
    }

    // exp(-x^2) with x split so that the rounding of x^2 does not show in the result.
    private static double ExpMinusSquare(double x)
    {
        var high = Math.Floor(x * 16) / 16;
        return Math.Exp(-high * high) * Math.Exp(-(x - high) * (x + high));
    }

    // exp(-x^2 / 2), split the same way.
    private static double ExpMinusHalfSquare(double x)
    {
        var high = Math.Floor(x * 16) / 16;
        return Math.Exp(-high * high / 2) * Math.Exp(-(x - high) * (x + high) / 2);
    }

    public static double NormalPdf(double z) => double.IsNaN(z) ? z : ExpMinusHalfSquare(Math.Abs(z)) / SqrtTwoPi;

    /// <summary>The standard normal distribution function, P(Z &lt;= z).</summary>
    public static double NormalCdf(double z)
    {
        if (double.IsNaN(z))
            return z;
        if (z > 1.4)
            return 1 - NormalLowerTail(-z);
        if (z >= -1.4)
            return 0.5 + 0.5 * ErfSeries(z / Math.Sqrt(2));
        return NormalLowerTail(z);
    }

    // P(Z <= z) for z < -1.4: 0.5 erfc(-z/sqrt 2), written in z so that no rounding of z/sqrt 2 shows.
    private static double NormalLowerTail(double z)
    {
        if (z < -38.5)
            return 0;
        return ExpMinusHalfSquare(-z) * -z / (SqrtTwoPi * TailFraction(z * z));
    }

    /// <summary>The standard normal quantile: z with P(Z &lt;= z) = p.</summary>
    public static double NormalQuantile(double p)
    {
        if (double.IsNaN(p) || p < 0 || p > 1)
            return double.NaN;
        if (p == 0)
            return double.NegativeInfinity;
        if (p == 1)
            return double.PositiveInfinity;

        // Work in the lower tail, where the probability has full relative precision; 1 - p is
        // exact for p >= 0.5.
        var q = p < 0.5 ? p : 1 - p;
        var x = AcklamQuantile(q);
        for (var i = 0; i < 3; i++)
        {
            var density = NormalPdf(x);
            if (density == 0)
                break;
            // Near the centre, compare P(Z <= x) - 1/2 with q - 1/2 (exact) so the small
            // difference keeps its relative precision.
            var error = x >= -1.4 ? 0.5 * ErfSeries(x / Math.Sqrt(2)) - (q - 0.5) : NormalLowerTail(x) - q;
            var u = error / density;
            if (!double.IsFinite(u))
                break;
            var step = u / (1 + x * u / 2);
            x -= step;
            if (Math.Abs(step) <= Math.Abs(x) * Epsilon)
                break;
        }

        return p < 0.5 ? x : -x;
    }

    // Acklam's rational approximation of the lower normal quantile for 0 < p <= 0.5 (relative
    // error about 1e-9); Halley steps on the exact distribution function finish the job.
    private static double AcklamQuantile(double p)
    {
        if (p < 0.02425)
        {
            var q = Math.Sqrt(-2 * Math.Log(p));
            return (((((-7.784894002430293e-03 * q - 3.223964580411365e-01) * q - 2.400758277161838e+00) * q
                       - 2.549732539343734e+00) * q + 4.374664141464968e+00) * q + 2.938163982698783e+00)
                   / ((((7.784695709041462e-03 * q + 3.224671290700398e-01) * q + 2.445134137142996e+00) * q
                       + 3.754408661907416e+00) * q + 1);
        }

        var r0 = p - 0.5;
        var r = r0 * r0;
        return (((((-3.969683028665376e+01 * r + 2.209460984245205e+02) * r - 2.759285104469687e+02) * r
                   + 1.383577518672690e+02) * r - 3.066479806614716e+01) * r + 2.506628277459239e+00) * r0
               / (((((-5.447609879822406e+01 * r + 1.615858368580409e+02) * r - 1.556989798598866e+02) * r
                   + 6.680131188771972e+01) * r - 1.328068155288572e+01) * r + 1);
    }

    // ---------------------------------------------------------------------------------------
    // Gamma function

    /// <summary>log Gamma(x) for x &gt; 0; NaN otherwise.</summary>
    public static double LogGamma(double x)
    {
        if (double.IsNaN(x) || x <= 0)
            return double.NaN;
        if (double.IsPositiveInfinity(x))
            return x;
        if (x < 0.5)
            return LogGamma1p(x) - Math.Log(x);
        if (x <= 1.5)
            return LogGamma1p(x - 1);
        if (x <= 2.5)
            return LogGamma1p(x - 2) + Log1p(x - 2);
        if (x < 15)
        {
            // Gamma(x) = Gamma(1 + z) * (z + 1)(z + 2)...(z + m) with z in [-0.5, 0.5).
            var m = (int)Math.Floor(x - 0.5);
            var z = x - 1 - m;
            var product = 1.0;
            for (var j = 1; j <= m; j++)
                product *= z + j;
            return LogGamma1p(z) + Math.Log(product);
        }

        return (x - 0.5) * Math.Log(x) - x + LogSqrtTwoPi + StirlingSeries(x);
    }

    /// <summary>Gamma(x); NaN at zero and the negative integers, infinity on overflow.</summary>
    public static double Gamma(double x)
    {
        if (double.IsNaN(x))
            return x;
        if (x <= 0)
        {
            if (x == Math.Floor(x))
                return double.NaN;
            // Reflection: Gamma(x) Gamma(1 - x) = pi / sin(pi x).
            var g = Gamma(1 - x);
            return double.IsInfinity(g) ? 0 : Math.PI / (SinPi(x) * g);
        }

        if (x == Math.Floor(x) && x <= 171)
        {
            var factorial = 1.0;
            for (var k = 2; k < x; k++)
                factorial *= k;
            return factorial;
        }

        if (x < 15)
        {
            if (x < 0.5)
                return Math.Exp(LogGamma1p(x)) / x;
            var m = (int)Math.Floor(x - 0.5);
            var z = x - 1 - m;
            var product = 1.0;
            for (var j = 1; j <= m; j++)
                product *= z + j;
            return Math.Exp(LogGamma1p(z)) * product;
        }

        if (x > 171.7)
            return double.PositiveInfinity;
        // x^(x - 1/2) e^-x in two halves so the power does not overflow before the exponential.
        var power = Math.Pow(x, (x - 0.5) / 2);
        return power * (power * Math.Exp(-x)) * SqrtTwoPi * Math.Exp(StirlingSeries(x));
    }

    // log Gamma(1 + z) for |z| <= 0.5: -gamma z + sum (-1)^k zeta(k) z^k / k. With zeta(k) split as
    // 1 + (zeta(k) - 1), the ones sum to z - log(1 + z) and the rest converge like 2^-k.
    private static double LogGamma1p(double z)
    {
        var sum = 0.0;
        var power = -z;
        for (var k = 2; k < ZetaTerms; k++)
        {
            power *= -z;
            var term = ZetaMinusOne[k] * power / k;
            sum += term;
            if (Math.Abs(term) <= Math.Abs(sum) * 1e-18)
                break;
        }

        return -EulerGamma * z + sum - Log1pmx(z);
    }

    // log Gamma(x) - [(x - 1/2) log x - x + log sqrt(2 pi)], the Stirling series, for x >= 10.
    private static double StirlingSeries(double x)
    {
        var inverse = 1 / x;
        var inverse2 = inverse * inverse;
        return inverse * (1.0 / 12 + inverse2 * (-1.0 / 360 + inverse2 * (1.0 / 1260 + inverse2 * (-1.0 / 1680
            + inverse2 * (1.0 / 1188 + inverse2 * (-691.0 / 360360 + inverse2 * (1.0 / 156 + inverse2 * (-3617.0 / 122400))))))));
    }

    /// <summary>
    /// The error of Stirling's formula, log Gamma(x + 1) - [(x + 1/2) log x - x + log sqrt(2 pi)]
    /// (Loader's stirlerr), for x &gt; 0.
    /// </summary>
    private static double StirlingError(double x)
    {
        if (x >= 10)
            return StirlingSeries(x);
        return LogGamma(x) - (x - 0.5) * Math.Log(x) + x - LogSqrtTwoPi;
    }

    // zeta(k) - 1 by direct summation of the first terms and Euler-Maclaurin for the tail.
    private static double[] BuildZetaMinusOne()
    {
        // Bernoulli numbers B2, B4, ..., B14.
        ReadOnlySpan<double> bernoulli = [1.0 / 6, -1.0 / 30, 1.0 / 42, -1.0 / 30, 5.0 / 66, -691.0 / 2730, 7.0 / 6];
        const int n = 10;
        var result = new double[ZetaTerms];
        for (var k = 2; k < ZetaTerms; k++)
        {
            var sum = 0.0;
            for (var i = n - 1; i >= 2; i--)
                sum += Math.Pow(i, -k);

            var tail = Math.Pow(n, 1 - k) / (k - 1) + 0.5 * Math.Pow(n, -k);
            var factorial = 1.0;
            var rising = 1.0;
            for (var j = 1; j <= bernoulli.Length; j++)
            {
                factorial *= (2.0 * j - 1) * (2.0 * j);
                rising = j == 1 ? k : rising * (k + 2.0 * j - 3) * (k + 2.0 * j - 2);
                tail += bernoulli[j - 1] / factorial * rising * Math.Pow(n, -k - 2 * j + 1);
            }

            result[k] = sum + tail;
        }

        return result;
    }

    /// <summary>log B(a, b) for a, b &gt; 0.</summary>
    public static double LogBeta(double a, double b)
    {
        var small = Math.Min(a, b);
        var large = Math.Max(a, b);
        if (small >= 10)
        {
            // Stirling's formula for all three gammas; the large terms are combined before they cancel.
            return LogSqrtTwoPi - 0.5 * Math.Log(large) + (small - 0.5) * Math.Log(small / (small + large))
                   + large * -Log1p(small / large) + StirlingSeries(small) + StirlingSeries(large) - StirlingSeries(small + large);
        }

        if (large >= 10)
        {
            // log Gamma(large) - log Gamma(large + small) by Stirling's formula.
            var difference = (large - 0.5) * -Log1p(small / large) - small * Math.Log(large + small) + small
                             + StirlingSeries(large) - StirlingSeries(large + small);
            return LogGamma(small) + difference;
        }

        return LogGamma(a) + LogGamma(b) - LogGamma(a + b);
    }

    // ---------------------------------------------------------------------------------------
    // Saddle-point densities (Loader, "Fast and accurate computation of binomial probabilities")

    /// <summary>x log(x / m) + m - x, the deviance term, without cancellation when x is near m.</summary>
    private static double Deviance(double x, double m)
    {
        if (x == 0)
            return m;
        if (Math.Abs(x - m) < 0.1 * (x + m))
        {
            var v = (x - m) / (x + m);
            var s = (x - m) * v;
            var term = 2 * x * v;
            var v2 = v * v;
            for (var j = 1; j < 1000; j++)
            {
                term *= v2;
                var next = s + term / (2 * j + 1);
                if (next == s)
                    return next;
                s = next;
            }

            return s;
        }

        return x * Math.Log(x / m) + m - x;
    }

    /// <summary>
    /// The binomial probability Gamma(n + 1) / (Gamma(x + 1) Gamma(n - x + 1)) p^x q^(n - x), for real
    /// 0 &lt;= x &lt;= n, with q = 1 - p given separately for precision.
    /// </summary>
    public static double BinomialTerm(double x, double n, double p, double q)
    {
        if (p == 0)
            return x == 0 ? 1 : 0;
        if (q == 0)
            return x == n ? 1 : 0;
        if (x == 0)
        {
            if (n == 0)
                return 1;
            return Math.Exp(p < 0.1 ? -Deviance(n, n * q) - n * p : n * Math.Log(q));
        }

        if (x == n)
            return Math.Exp(q < 0.1 ? -Deviance(n, n * p) - n * q : n * Math.Log(p));
        if (x < 0 || x > n)
            return 0;
        return BinomialTermCore(x, n - x, p, q);
    }

    // The binomial term for x and r = n - x, both positive and given separately: when one is much
    // smaller than n, recovering it as n - x would lose its precision.
    private static double BinomialTermCore(double x, double r, double p, double q)
    {
        var n = x + r;
        var lc = StirlingError(n) - StirlingError(x) - StirlingError(r) - Deviance(x, n * p) - Deviance(r, n * q);
        var lf = LogTwoPi + Math.Log(x) + Math.Log(r) - Math.Log(n);
        return Math.Exp(lc - 0.5 * lf);
    }

    /// <summary>The Poisson probability m^x e^-m / Gamma(x + 1) for real x &gt;= 0.</summary>
    public static double PoissonTerm(double x, double m)
    {
        if (m == 0)
            return x == 0 ? 1 : 0;
        if (x < 0)
            return 0;
        if (x == 0)
            return Math.Exp(-m);
        if (double.IsInfinity(m))
            return 0;
        return Math.Exp(-StirlingError(x) - Deviance(x, m)) / Math.Sqrt(2 * Math.PI * x);
    }

    /// <summary>The density of Student's t distribution with n degrees of freedom (Loader's form, R's dt).</summary>
    public static double StudentDensity(double x, double n)
    {
        var t = -Deviance(n / 2, (n + 1) / 2) + StirlingError((n + 1) / 2) - StirlingError(n / 2);
        var x2n = x * x / n;
        double logRoot, u;
        var large = x2n > 1 / Epsilon;
        if (large)
        {
            // log sqrt(1 + x^2/n) is log|x| - log sqrt(n) to working precision.
            logRoot = Math.Log(Math.Abs(x)) - Math.Log(n) / 2;
            u = n * logRoot;
        }
        else if (x2n > 0.2)
        {
            logRoot = Math.Log(1 + x2n) / 2;
            u = n * logRoot;
        }
        else
        {
            logRoot = Log1p(x2n) / 2;
            u = -Deviance(n / 2, (n + x * x) / 2) + x * x / 2;
        }

        var inverseRoot = large ? Math.Sqrt(n) / Math.Abs(x) : Math.Exp(-logRoot);
        return Math.Exp(t - u) / SqrtTwoPi * inverseRoot;
    }

    // ---------------------------------------------------------------------------------------
    // Regularized incomplete gamma function

    /// <summary>P(a, x) = gamma(a, x) / Gamma(a), the lower regularized incomplete gamma function.</summary>
    public static double GammaP(double a, double x) => IncompleteGamma(a, x, upper: false);

    /// <summary>Q(a, x) = 1 - P(a, x).</summary>
    public static double GammaQ(double a, double x) => IncompleteGamma(a, x, upper: true);

    private static double IncompleteGamma(double a, double x, bool upper)
    {
        if (double.IsNaN(a) || double.IsNaN(x) || a <= 0)
            return double.NaN;
        if (x <= 0)
            return upper ? 1 : 0;
        if (double.IsPositiveInfinity(x))
            return upper ? 0 : 1;

        if (x < a + 1)
        {
            if (upper && a < 1)
                return SmallShapeUpperGamma(a, x);

            // P = x^a e^-x / Gamma(a + 1) * sum x^n / ((a + 1)...(a + n)).
            var sum = 1.0;
            var term = 1.0;
            var ap = a;
            for (var n = 0; n < MaxIterations; n++)
            {
                ap += 1;
                term *= x / ap;
                sum += term;
                if (term <= sum * Epsilon)
                    break;
            }

            var p = PoissonTerm(a, x) * sum;
            return upper ? 1 - p : p;
        }

        // Q = x^a e^-x / Gamma(a) / (x + 1 - a - 1(1 - a)/(x + 3 - a - 2(2 - a)/(x + 5 - a - ...))).
        var b = x + 1 - a;
        var c = 1 / Tiny;
        var d = 1 / b;
        var h = d;
        for (var i = 1; i < MaxIterations; i++)
        {
            var an = -i * (i - a);
            b += 2;
            d = an * d + b;
            if (Math.Abs(d) < Tiny)
                d = Tiny;
            c = b + an / c;
            if (Math.Abs(c) < Tiny)
                c = Tiny;
            d = 1 / d;
            var delta = d * c;
            h *= delta;
            if (Math.Abs(delta - 1) <= Epsilon)
                break;
        }

        var q = a * PoissonTerm(a, x) * h;
        return upper ? q : 1 - q;
    }

    // Q(a, x) for a < 1 and x < a + 1, where 1 - P would cancel: from P = x^a / Gamma(a + 1) *
    // (1 + a sum_{n>=1} (-x)^n / (n! (a + n))), Q = -expm1(a log x - log Gamma(a + 1)) - x^a / Gamma(a + 1) * a * sum.
    private static double SmallShapeUpperGamma(double a, double x)
    {
        var logFront = a * Math.Log(x) - (a < 0.5 ? LogGamma1p(a) : LogGamma(1 + a));
        var sum = 0.0;
        var term = 1.0;
        for (var n = 1; n < 1000; n++)
        {
            term *= -x / n;
            var part = term / (a + n);
            sum += part;
            if (Math.Abs(part) <= Math.Abs(sum) * Epsilon)
                break;
        }

        return -Expm1(logFront) - Math.Exp(logFront) * a * sum;
    }

    /// <summary>The density x^(a-1) e^-x / Gamma(a) of the standard gamma distribution.</summary>
    public static double GammaDensity(double a, double x)
    {
        if (x < 0)
            return 0;
        if (x == 0)
            return a < 1 ? double.PositiveInfinity : a == 1 ? 1 : 0;
        return a < 1 ? PoissonTerm(a, x) * a / x : PoissonTerm(a - 1, x);
    }

    /// <summary>x with P(a, x) = p; q = 1 - p is passed separately so the upper tail keeps its precision.</summary>
    public static double InverseGammaP(double a, double p, double q)
    {
        if (double.IsNaN(a) || a <= 0 || double.IsNaN(p) || double.IsNaN(q) || p < 0 || q < 0)
            return double.NaN;
        if (p == 0)
            return 0;
        if (q == 0)
            return double.PositiveInfinity;

        // Starting guess (Numerical Recipes): Wilson-Hilferty for a > 1, the small-x form otherwise.
        double x;
        if (a > 1)
        {
            var t = Math.Sqrt(-2 * Math.Log(Math.Min(p, q)));
            var z = (2.30753 + t * 0.27061) / (1 + t * (0.99229 + t * 0.04481)) - t;
            if (p < 0.5)
                z = -z;
            x = Math.Max(1e-3, a * Math.Pow(1 - 1 / (9 * a) - z / (3 * Math.Sqrt(a)), 3));
        }
        else
        {
            var t = 1 - a * (0.253 + a * 0.12);
            x = p < t ? Math.Pow(p / t, 1 / a) : 1 - Math.Log(q / (1 - t));
        }

        // The answer is below the smallest double.
        if (x == 0)
            return 0;
        if (!(x > 0) || double.IsInfinity(x))
            x = a;

        return p <= q
            ? SolveTail(v => GammaP(a, v), v => GammaDensity(a, v), increasing: true, p, x, double.PositiveInfinity)
            : SolveTail(v => GammaQ(a, v), v => GammaDensity(a, v), increasing: false, q, x, double.PositiveInfinity);
    }

    /// <summary>
    /// Solves tail(x) = target for a monotone tail probability by Newton's method on log(tail)
    /// against log(x), which is nearly linear in the far tails, kept inside a shrinking bracket
    /// [0, limit]. Callers solve for the smaller of the two tails, so the target has full relative
    /// precision.
    /// </summary>
    private static double SolveTail(Func<double, double> tail, Func<double, double> density, bool increasing,
        double target, double x, double limit)
    {
        var logTarget = Math.Log(target);
        double low = 0, high = limit;
        for (var i = 0; i < 300; i++)
        {
            var t = tail(x);
            if (t == target)
                return x;
            if ((t > target) == increasing)
                high = x;
            else
                low = x;

            var next = double.NaN;
            if (t > 0)
            {
                var d = density(x);
                if (d > 0 && double.IsFinite(d))
                {
                    var exponent = (Math.Log(t) - logTarget) * t / (x * d);
                    if (increasing)
                        exponent = -exponent;
                    if (Math.Abs(exponent) < 50)
                        next = x * Math.Exp(exponent);
                }
            }

            if (Math.Abs(next - x) <= 4 * Epsilon * x)
                return next;
            if (!(next > low && next < high))
            {
                next = double.IsInfinity(high) ? 4 * Math.Max(Math.Max(x, low), 1)
                    : low == 0 ? high / 16
                    : high > 4 * low ? Math.Sqrt(low * high)
                    : (low + high) / 2;
            }

            if (next == x)
                return x;
            x = next;
        }

        return x;
    }

    // ---------------------------------------------------------------------------------------
    // Regularized incomplete beta function

    /// <summary>x^a y^b / B(a, b) with y = 1 - x.</summary>
    private static double BetaPrefix(double a, double b, double x, double y) =>
        BinomialTermCore(a, b, x, y) * (a * b / (a + b));

    /// <summary>
    /// I_x(a, b), the regularized incomplete beta function, or 1 - I_x(a, b) when
    /// <paramref name="upper"/>; y = 1 - x is passed separately so values near 1 keep their precision.
    /// </summary>
    public static double BetaRegularized(double a, double b, double x, double y, bool upper = false)
    {
        if (double.IsNaN(a) || double.IsNaN(b) || double.IsNaN(x) || a <= 0 || b <= 0)
            return double.NaN;
        if (x <= 0)
            return upper ? 1 : 0;
        if (y <= 0)
            return upper ? 0 : 1;

        // The continued fraction converges fast below the mean; above it, use I_x(a,b) = 1 - I_y(b,a).
        if (x > (a + 1) / (a + b + 2))
        {
            (a, b) = (b, a);
            (x, y) = (y, x);
            upper = !upper;
        }

        var result = BetaPrefix(a, b, x, y) / a * BetaFraction(a, b, x, y);
        if (!upper)
        {
            // With x near 1 (a much larger than b) the fraction's terms cancel and lose about
            // eps / y; going through the other side costs only eps / result, so use it when that
            // is clearly smaller. Beyond its own mean the other fraction converges slowly, so only
            // when the result is not small. This is Student's t with many degrees of freedom.
            if (x > 0.5 && result > 1e-3 && y < 0.01 * result)
                return 1 - BetaPrefix(b, a, y, x) / b * BetaFraction(b, a, y, x);
            return result;
        }
        if (result <= 0.5)
            return 1 - result;

        // The wanted tail is the small one and 1 - result would cancel: evaluate it directly, by the
        // continued fraction on the other side (slower there, but convergent).
        return BetaPrefix(b, a, y, x) / b * BetaFraction(b, a, y, x);
    }

    // The continued fraction of I_x(a, b) (Numerical Recipes' betacf, modified Lentz).
    private static double BetaFraction(double a, double b, double x, double y)
    {
        var qab = a + b;
        var qap = a + 1;
        var qam = a - 1;
        var c = 1.0;
        // 1 - (a + b) x / (a + 1), written with y so it does not cancel for x near 1.
        var d = (a * y - b * x + 1) / qap;
        if (Math.Abs(d) < Tiny)
            d = Tiny;
        d = 1 / d;
        var h = d;
        for (var m = 1; m < MaxIterations; m++)
        {
            var m2 = 2.0 * m;
            var aa = m * (b - m) * x / ((qam + m2) * (a + m2));
            d = 1 + aa * d;
            if (Math.Abs(d) < Tiny)
                d = Tiny;
            c = 1 + aa / c;
            if (Math.Abs(c) < Tiny)
                c = Tiny;
            d = 1 / d;
            h *= d * c;

            aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
            d = 1 + aa * d;
            if (Math.Abs(d) < Tiny)
                d = Tiny;
            c = 1 + aa / c;
            if (Math.Abs(c) < Tiny)
                c = Tiny;
            d = 1 / d;
            var delta = d * c;
            h *= delta;
            if (Math.Abs(delta - 1) <= Epsilon)
                break;
        }

        return h;
    }

    /// <summary>The beta density x^(a-1) y^(b-1) / B(a, b), y = 1 - x.</summary>
    public static double BetaDensity(double a, double b, double x, double y)
    {
        if (x < 0 || y < 0)
            return 0;
        if (x == 0)
            return a < 1 ? double.PositiveInfinity : a == 1 ? b : 0;
        if (y == 0)
            return b < 1 ? double.PositiveInfinity : b == 1 ? a : 0;
        return BetaPrefix(a, b, x, y) / (x * y);
    }

    /// <summary>
    /// x with I_x(a, b) = p, returned together with 1 - x; q = 1 - p is passed separately so that
    /// both tails keep their precision.
    /// </summary>
    public static (double X, double Y) InverseBeta(double a, double b, double p, double q)
    {
        if (double.IsNaN(a) || double.IsNaN(b) || a <= 0 || b <= 0 || double.IsNaN(p) || double.IsNaN(q) || p < 0 || q < 0)
            return (double.NaN, double.NaN);
        if (p == 0)
            return (0, 1);
        if (q == 0)
            return (1, 0);

        var x = BetaGuess(a, b, p, q);

        // Iterate on the smaller of x and 1 - x, so that it keeps its relative precision: if the
        // guess is above one half, solve I_y(b, a) = q instead.
        var flipped = x > 0.5;
        if (flipped)
        {
            (a, b) = (b, a);
            (p, q) = (q, p);
            x = 1 - x;
        }

        if (!(x > 0 && x < 1))
            x = 0.5;

        var alpha = a;
        var beta = b;
        x = p <= q
            ? SolveTail(v => BetaRegularized(alpha, beta, v, 1 - v), v => BetaDensity(alpha, beta, v, 1 - v), increasing: true, p, x, 1)
            : SolveTail(v => BetaRegularized(alpha, beta, v, 1 - v, upper: true), v => BetaDensity(alpha, beta, v, 1 - v), increasing: false, q, x, 1);
        return flipped ? (1 - x, x) : (x, 1 - x);
    }

    // Starting guess for the inverse beta function (Numerical Recipes' invbetai).
    private static double BetaGuess(double a, double b, double p, double q)
    {
        if (a >= 1 && b >= 1)
        {
            var t = Math.Sqrt(-2 * Math.Log(Math.Min(p, q)));
            var z = (2.30753 + t * 0.27061) / (1 + t * (0.99229 + t * 0.04481)) - t;
            if (p < 0.5)
                z = -z;
            var al = (z * z - 3) / 6;
            var h = 2 / (1 / (2 * a - 1) + 1 / (2 * b - 1));
            var w = z * Math.Sqrt(al + h) / h - (1 / (2 * b - 1) - 1 / (2 * a - 1)) * (al + 5.0 / 6 - 2 / (3 * h));
            return a / (a + b * Math.Exp(2 * w));
        }

        var lna = Math.Log(a / (a + b));
        var lnb = Math.Log(b / (a + b));
        var ta = Math.Exp(a * lna) / a;
        var tb = Math.Exp(b * lnb) / b;
        var sum = ta + tb;
        return p < ta / sum ? Math.Pow(a * sum * p, 1 / a) : 1 - Math.Pow(b * sum * q, 1 / b);
    }
}
