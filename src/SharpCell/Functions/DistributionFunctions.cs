using System;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// Reads the scalar arguments of a distribution function in order and keeps the first one that
/// does not convert: Excel reports the error of the leftmost bad argument, and only then checks
/// the values against the function's domain.
/// </summary>
internal sealed class DistributionArguments(FunctionCall call)
{
    private CellValue _error;

    public bool Failed => _error.IsError;

    public CellValue Error => _error;

    public double Number(int index) => Take(call.Number(index));

    public double Number(int index, double absent) => Take(call.Number(index, absent));

    /// <summary>Degrees of freedom and the like: Excel truncates these toward zero.</summary>
    public double Integer(int index) => Take(call.Integer(index));

    /// <summary>
    /// A count of events: truncated like <see cref="Integer"/>, but a negative value stays negative
    /// so that it fails the domain check (POISSON.DIST(-0.000001, 5) is #NUM!, not P(X = 0)).
    /// </summary>
    public double Count(int index)
    {
        var value = Take(call.Number(index));
        return value < 0 ? Math.Min(Math.Truncate(value), -1) : Math.Truncate(value);
    }

    public bool Boolean(int index)
    {
        var value = call.Boolean(index);
        if (Failed)
            return false;
        if (value.IsError)
        {
            _error = value;
            return false;
        }

        return value.AsBoolean();
    }

    /// <summary>A number for the engineering functions (ERF, ERFC): numeric text converts, logical values do not.</summary>
    public double StrictNumber(int index)
    {
        var value = call.Value(index);
        return Take(value.Kind == CellValueKind.Boolean ? CellValue.Error(ErrorKind.Value) : call.Number(index));
    }

    private double Take(CellValue value)
    {
        if (Failed)
            return double.NaN;
        if (value.IsError)
        {
            _error = value;
            return double.NaN;
        }

        return value.AsNumber();
    }
}

/// <summary>
/// Continuous distributions (normal, lognormal, Student's t, chi-square, F, gamma, beta,
/// exponential, Weibull), their inverses, and the special functions Excel exposes directly
/// (ERF, GAMMA, FISHER...). The numerics live in <see cref="Special"/>.
/// </summary>
internal static class DistributionFunctions
{
    private static readonly ArgumentKind[] Scalars = [ArgumentKind.Value];

    // CHIDIST, CHISQ.DIST and friends reject more than 10^10 degrees of freedom.
    private const double MaxChiSquareFreedom = 1e10;

    public static void Register(FunctionRegistry registry)
    {
        Add(registry, "NORM.DIST", 4, 4, NormDist);
        Add(registry, "NORMDIST", 4, 4, NormDist);
        Add(registry, "NORM.INV", 3, 3, NormInv);
        Add(registry, "NORMINV", 3, 3, NormInv);
        Add(registry, "NORM.S.DIST", 2, 2, call => StandardNormalDist(call, legacy: false));
        Add(registry, "NORMSDIST", 1, 1, call => StandardNormalDist(call, legacy: true));
        Add(registry, "NORM.S.INV", 1, 1, StandardNormalInv);
        Add(registry, "NORMSINV", 1, 1, StandardNormalInv);
        Add(registry, "LOGNORM.DIST", 4, 4, call => LognormDist(call, legacy: false));
        Add(registry, "LOGNORMDIST", 3, 3, call => LognormDist(call, legacy: true));
        Add(registry, "LOGNORM.INV", 3, 3, LognormInv);
        Add(registry, "LOGINV", 3, 3, LognormInv);
        Add(registry, "STANDARDIZE", 3, 3, Standardize);
        Add(registry, "PHI", 1, 1, Phi);
        Add(registry, "GAUSS", 1, 1, Gauss);
        Add(registry, "CONFIDENCE.NORM", 3, 3, ConfidenceNorm);
        Add(registry, "CONFIDENCE", 3, 3, ConfidenceNorm);
        Add(registry, "CONFIDENCE.T", 3, 3, ConfidenceT);

        Add(registry, "EXPON.DIST", 3, 3, ExponDist);
        Add(registry, "EXPONDIST", 3, 3, ExponDist);
        Add(registry, "WEIBULL.DIST", 4, 4, WeibullDist);
        Add(registry, "WEIBULL", 4, 4, WeibullDist);
        Add(registry, "GAMMA.DIST", 4, 4, GammaDist);
        Add(registry, "GAMMADIST", 4, 4, GammaDist);
        Add(registry, "GAMMA.INV", 3, 3, GammaInv);
        Add(registry, "GAMMAINV", 3, 3, GammaInv);
        Add(registry, "GAMMA", 1, 1, GammaFunction);
        Add(registry, "GAMMALN", 1, 1, GammaLn);
        Add(registry, "GAMMALN.PRECISE", 1, 1, GammaLn);
        Add(registry, "BETA.DIST", 4, 6, call => BetaDist(call, legacy: false));
        Add(registry, "BETADIST", 3, 5, call => BetaDist(call, legacy: true));
        Add(registry, "BETA.INV", 3, 5, BetaInv);
        Add(registry, "BETAINV", 3, 5, BetaInv);

        Add(registry, "CHISQ.DIST", 3, 3, ChiSquareDist);
        Add(registry, "CHISQ.DIST.RT", 2, 2, ChiSquareRightTail);
        Add(registry, "CHIDIST", 2, 2, ChiSquareRightTail);
        Add(registry, "CHISQ.INV", 2, 2, call => ChiSquareInv(call, rightTail: false));
        Add(registry, "CHISQ.INV.RT", 2, 2, call => ChiSquareInv(call, rightTail: true));
        Add(registry, "CHIINV", 2, 2, call => ChiSquareInv(call, rightTail: true));
        Add(registry, "F.DIST", 4, 4, FDist);
        Add(registry, "F.DIST.RT", 3, 3, FRightTail);
        Add(registry, "FDIST", 3, 3, FRightTail);
        Add(registry, "F.INV", 3, 3, call => FInv(call, rightTail: false));
        Add(registry, "F.INV.RT", 3, 3, call => FInv(call, rightTail: true));
        Add(registry, "FINV", 3, 3, call => FInv(call, rightTail: true));
        Add(registry, "T.DIST", 3, 3, TDist);
        Add(registry, "T.DIST.2T", 2, 2, TDistTwoTailed);
        Add(registry, "T.DIST.RT", 2, 2, TDistRightTail);
        Add(registry, "TDIST", 3, 3, TDistLegacy);
        Add(registry, "T.INV", 2, 2, TInv);
        Add(registry, "T.INV.2T", 2, 2, TInvTwoTailed);
        Add(registry, "TINV", 2, 2, TInvTwoTailed);

        Add(registry, "FISHER", 1, 1, Fisher);
        Add(registry, "FISHERINV", 1, 1, FisherInv);
        Add(registry, "ERF", 1, 2, Erf);
        Add(registry, "ERF.PRECISE", 1, 1, call => ErrorFunction(call, complement: false));
        Add(registry, "ERFC", 1, 1, call => ErrorFunction(call, complement: true));
        Add(registry, "ERFC.PRECISE", 1, 1, call => ErrorFunction(call, complement: true));
    }

    private static void Add(FunctionRegistry registry, string name, int min, int max, FunctionBody body) =>
        registry.Add(new FunctionInfo(name, min, max, Scalars, body));

    // -----------------------------------------------------------------------------------------
    // Normal and lognormal

    private static Operand NormDist(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(0);
        var mean = args.Number(1);
        var sd = args.Number(2);
        var cumulative = args.Boolean(3);
        if (args.Failed)
            return args.Error;
        if (sd <= 0)
            return CellValue.Error(ErrorKind.Num);

        var z = (x - mean) / sd;
        return CellValue.Number(cumulative ? Special.NormalCdf(z) : Special.NormalPdf(z) / sd);
    }

    private static Operand NormInv(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var p = args.Number(0);
        var mean = args.Number(1);
        var sd = args.Number(2);
        if (args.Failed)
            return args.Error;
        if (p <= 0 || p >= 1 || sd <= 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(mean + sd * Special.NormalQuantile(p));
    }

    // NORMSDIST has no cumulative argument: it is always the distribution function.
    private static Operand StandardNormalDist(FunctionCall call, bool legacy)
    {
        var args = new DistributionArguments(call);
        var z = args.Number(0);
        var cumulative = legacy || args.Boolean(1);
        if (args.Failed)
            return args.Error;
        return CellValue.Number(cumulative ? Special.NormalCdf(z) : Special.NormalPdf(z));
    }

    private static Operand StandardNormalInv(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var p = args.Number(0);
        if (args.Failed)
            return args.Error;
        if (p <= 0 || p >= 1)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(Special.NormalQuantile(p));
    }

    private static Operand LognormDist(FunctionCall call, bool legacy)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(0);
        var mean = args.Number(1);
        var sd = args.Number(2);
        var cumulative = legacy || args.Boolean(3);
        if (args.Failed)
            return args.Error;
        if (x <= 0 || sd <= 0)
            return CellValue.Error(ErrorKind.Num);

        var z = (Math.Log(x) - mean) / sd;
        return CellValue.Number(cumulative ? Special.NormalCdf(z) : Special.NormalPdf(z) / (x * sd));
    }

    private static Operand LognormInv(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var p = args.Number(0);
        var mean = args.Number(1);
        var sd = args.Number(2);
        if (args.Failed)
            return args.Error;
        if (p <= 0 || p >= 1 || sd <= 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(Math.Exp(mean + sd * Special.NormalQuantile(p)));
    }

    private static Operand Standardize(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(0);
        var mean = args.Number(1);
        var sd = args.Number(2);
        if (args.Failed)
            return args.Error;
        if (sd <= 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number((x - mean) / sd);
    }

    private static Operand Phi(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(0);
        return args.Failed ? args.Error : CellValue.Number(Special.NormalPdf(x));
    }

    // Excel subtracts one half from the rounded distribution function; GAUSS(1e-7) shows it.
    private static Operand Gauss(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var z = args.Number(0);
        return args.Failed ? args.Error : CellValue.Number(Special.NormalCdf(z) - 0.5);
    }

    private static Operand ConfidenceNorm(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var alpha = args.Number(0);
        var sd = args.Number(1);
        var size = args.Integer(2);
        if (args.Failed)
            return args.Error;
        if (alpha <= 0 || alpha >= 1 || sd <= 0 || size < 1)
            return CellValue.Error(ErrorKind.Num);

        // The upper alpha/2 quantile, taken from the lower tail where alpha/2 is exact.
        return CellValue.Number(-Special.NormalQuantile(alpha / 2) * sd / Math.Sqrt(size));
    }

    private static Operand ConfidenceT(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var alpha = args.Number(0);
        var sd = args.Number(1);
        var size = args.Integer(2);
        if (args.Failed)
            return args.Error;
        if (alpha <= 0 || alpha >= 1 || sd <= 0 || size < 1)
            return CellValue.Error(ErrorKind.Num);
        if (size == 1)
            return CellValue.Error(ErrorKind.Div0);
        return CellValue.Number(StudentTwoTailedQuantile(alpha, size - 1) * sd / Math.Sqrt(size));
    }

    // -----------------------------------------------------------------------------------------
    // Exponential, Weibull, gamma, beta

    private static Operand ExponDist(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(0);
        var lambda = args.Number(1);
        var cumulative = args.Boolean(2);
        if (args.Failed)
            return args.Error;
        if (x < 0 || lambda <= 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(cumulative ? -Special.Expm1(-lambda * x) : lambda * Math.Exp(-lambda * x));
    }

    private static Operand WeibullDist(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(0);
        var alpha = args.Number(1);
        var beta = args.Number(2);
        var cumulative = args.Boolean(3);
        if (args.Failed)
            return args.Error;
        if (x < 0 || alpha <= 0 || beta <= 0)
            return CellValue.Error(ErrorKind.Num);

        var power = Math.Pow(x / beta, alpha);
        if (cumulative)
            return CellValue.Number(-Special.Expm1(-power));
        if (x == 0)
            return alpha < 1 ? CellValue.Error(ErrorKind.Num) : CellValue.Number(alpha == 1 ? 1 / beta : 0);
        return CellValue.Number(alpha / beta * Math.Pow(x / beta, alpha - 1) * Math.Exp(-power));
    }

    private static Operand GammaDist(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(0);
        var alpha = args.Number(1);
        var beta = args.Number(2);
        var cumulative = args.Boolean(3);
        if (args.Failed)
            return args.Error;
        if (x < 0 || alpha <= 0 || beta <= 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(cumulative ? Special.GammaP(alpha, x / beta) : Special.GammaDensity(alpha, x / beta) / beta);
    }

    private static Operand GammaInv(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var p = args.Number(0);
        var alpha = args.Number(1);
        var beta = args.Number(2);
        if (args.Failed)
            return args.Error;
        if (p < 0 || p >= 1 || alpha <= 0 || beta <= 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(beta * Special.InverseGammaP(alpha, p, 1 - p));
    }

    private static Operand GammaFunction(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(0);
        return args.Failed ? args.Error : CellValue.Number(Special.Gamma(x));
    }

    private static Operand GammaLn(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(0);
        if (args.Failed)
            return args.Error;
        return x <= 0 ? CellValue.Error(ErrorKind.Num) : CellValue.Number(Special.LogGamma(x));
    }

    // BETA.DIST(x, alpha, beta, cumulative, [A], [B]); BETADIST has no cumulative argument.
    private static Operand BetaDist(FunctionCall call, bool legacy)
    {
        var args = new DistributionArguments(call);
        var bounds = legacy ? 3 : 4;
        var x = args.Number(0);
        var alpha = args.Number(1);
        var beta = args.Number(2);
        var cumulative = legacy || args.Boolean(3);
        var lower = args.Number(bounds, 0);
        var upper = args.Number(bounds + 1, 1);
        if (args.Failed)
            return args.Error;
        if (alpha <= 0 || beta <= 0 || lower >= upper || x < lower || x > upper)
            return CellValue.Error(ErrorKind.Num);

        var width = upper - lower;
        var t = (x - lower) / width;
        var s = (upper - x) / width;
        return CellValue.Number(cumulative ? Special.BetaRegularized(alpha, beta, t, s) : Special.BetaDensity(alpha, beta, t, s) / width);
    }

    private static Operand BetaInv(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var p = args.Number(0);
        var alpha = args.Number(1);
        var beta = args.Number(2);
        var lower = args.Number(3, 0);
        var upper = args.Number(4, 1);
        if (args.Failed)
            return args.Error;
        if (p <= 0 || p >= 1 || alpha <= 0 || beta <= 0 || lower >= upper)
            return CellValue.Error(ErrorKind.Num);

        var (t, _) = Special.InverseBeta(alpha, beta, p, 1 - p);
        return CellValue.Number(lower + (upper - lower) * t);
    }

    // -----------------------------------------------------------------------------------------
    // Chi-square, F, Student's t

    private static Operand ChiSquareDist(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(0);
        var freedom = args.Integer(1);
        var cumulative = args.Boolean(2);
        if (args.Failed)
            return args.Error;
        if (x < 0 || freedom < 1 || freedom > MaxChiSquareFreedom)
            return CellValue.Error(ErrorKind.Num);
        var shape = freedom / 2;
        return CellValue.Number(cumulative ? Special.GammaP(shape, x / 2) : Special.GammaDensity(shape, x / 2) / 2);
    }

    private static Operand ChiSquareRightTail(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(0);
        var freedom = args.Integer(1);
        if (args.Failed)
            return args.Error;
        if (x < 0 || freedom < 1 || freedom > MaxChiSquareFreedom)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(Special.GammaQ(freedom / 2, x / 2));
    }

    private static Operand ChiSquareInv(FunctionCall call, bool rightTail)
    {
        var args = new DistributionArguments(call);
        var p = args.Number(0);
        var freedom = args.Integer(1);
        if (args.Failed)
            return args.Error;
        if (p < 0 || p > 1 || freedom < 1 || freedom > MaxChiSquareFreedom)
            return CellValue.Error(ErrorKind.Num);

        var lowerTail = rightTail ? 1 - p : p;
        var upperTail = rightTail ? p : 1 - p;
        if (upperTail == 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(2 * Special.InverseGammaP(freedom / 2, lowerTail, upperTail));
    }

    private static bool TryReadF(DistributionArguments args, out double x, out double d1, out double d2)
    {
        x = args.Number(0);
        d1 = args.Integer(1);
        d2 = args.Integer(2);
        return !args.Failed;
    }

    private static Operand FDist(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        TryReadF(args, out var x, out var d1, out var d2);
        var cumulative = args.Boolean(3);
        if (args.Failed)
            return args.Error;
        if (x < 0 || d1 < 1 || d2 < 1)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(cumulative ? FisherCdf(x, d1, d2, upper: false) : FisherDensity(x, d1, d2));
    }

    private static Operand FRightTail(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        if (!TryReadF(args, out var x, out var d1, out var d2))
            return args.Error;
        if (x < 0 || d1 < 1 || d2 < 1)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(FisherCdf(x, d1, d2, upper: true));
    }

    private static Operand FInv(FunctionCall call, bool rightTail)
    {
        var args = new DistributionArguments(call);
        if (!TryReadF(args, out var p, out var d1, out var d2))
            return args.Error;
        if (p < 0 || p > 1 || d1 < 1 || d2 < 1)
            return CellValue.Error(ErrorKind.Num);

        var lowerTail = rightTail ? 1 - p : p;
        var upperTail = rightTail ? p : 1 - p;
        if (upperTail == 0)
            return CellValue.Error(ErrorKind.Num);
        if (lowerTail == 0)
            return CellValue.Number(0);
        var (t, s) = Special.InverseBeta(d1 / 2, d2 / 2, lowerTail, upperTail);
        return CellValue.Number(d2 * t / (d1 * s));
    }

    /// <summary>The F distribution function (or its upper tail) with d1 and d2 degrees of freedom.</summary>
    internal static double FisherCdf(double x, double d1, double d2, bool upper)
    {
        var denominator = d1 * x + d2;
        return Special.BetaRegularized(d1 / 2, d2 / 2, d1 * x / denominator, d2 / denominator, upper);
    }

    // The F density by Loader's binomial form (R's df), accurate in the tails.
    private static double FisherDensity(double x, double m, double n)
    {
        if (x == 0)
            return m > 2 ? 0 : m == 2 ? 1 : double.PositiveInfinity;
        var f = 1 / (n + x * m);
        var q = n * f;
        var p = x * m * f;
        if (m >= 2)
            return m * q / 2 * Special.BinomialTerm((m - 2) / 2, (m + n - 2) / 2, p, q);
        return m * m * q / (2 * p * (m + n)) * Special.BinomialTerm(m / 2, (m + n) / 2, p, q);
    }

    private static bool TryReadT(DistributionArguments args, out double x, out double freedom)
    {
        x = args.Number(0);
        freedom = args.Integer(1);
        return !args.Failed;
    }

    private static Operand TDist(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        TryReadT(args, out var x, out var freedom);
        var cumulative = args.Boolean(2);
        if (args.Failed)
            return args.Error;
        if (freedom < 1)
            return CellValue.Error(ErrorKind.Num);
        if (!cumulative)
            return CellValue.Number(Special.StudentDensity(x, freedom));
        return CellValue.Number(x < 0 ? StudentTail(x, freedom) : 1 - StudentTail(x, freedom));
    }

    private static Operand TDistTwoTailed(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        if (!TryReadT(args, out var x, out var freedom))
            return args.Error;
        if (x < 0 || freedom < 1)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(2 * StudentTail(x, freedom));
    }

    private static Operand TDistRightTail(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        if (!TryReadT(args, out var x, out var freedom))
            return args.Error;
        if (freedom < 1)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(x >= 0 ? StudentTail(x, freedom) : 1 - StudentTail(x, freedom));
    }

    // TDIST(x, degrees_freedom, tails): x must not be negative, tails is 1 or 2.
    private static Operand TDistLegacy(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        TryReadT(args, out var x, out var freedom);
        var tails = args.Integer(2);
        if (args.Failed)
            return args.Error;
        if (x < 0 || freedom < 1 || (tails != 1 && tails != 2))
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(tails * StudentTail(x, freedom));
    }

    private static Operand TInv(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        if (!TryReadT(args, out var p, out var freedom))
            return args.Error;
        if (p <= 0 || p >= 1 || freedom < 1)
            return CellValue.Error(ErrorKind.Num);
        if (p == 0.5)
            return CellValue.Number(0);

        // The one-tailed probability beyond |t| is the smaller tail, exact as 1 - p above one half.
        var tail = p < 0.5 ? p : 1 - p;
        var t = StudentTwoTailedQuantile(2 * tail, freedom);
        return CellValue.Number(p < 0.5 ? -t : t);
    }

    private static Operand TInvTwoTailed(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        if (!TryReadT(args, out var p, out var freedom))
            return args.Error;
        if (p <= 0 || p > 1 || freedom < 1)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(StudentTwoTailedQuantile(p, freedom));
    }

    /// <summary>P(T &gt; |t|) for Student's t with the given degrees of freedom.</summary>
    internal static double StudentTail(double t, double freedom)
    {
        // P(|T| > |t|) = I_x(n/2, 1/2) with x = n / (n + t^2); 1 - x is computed directly.
        var t2 = t * t;
        if (double.IsInfinity(t2))
            return 0;
        var denominator = freedom + t2;
        return 0.5 * Special.BetaRegularized(freedom / 2, 0.5, freedom / denominator, t2 / denominator);
    }

    /// <summary>|t| with P(|T| &gt; |t|) = p, for 0 &lt; p &lt;= 1.</summary>
    internal static double StudentTwoTailedQuantile(double p, double freedom)
    {
        if (p >= 1)
            return 0;
        var (x, y) = Special.InverseBeta(freedom / 2, 0.5, p, 1 - p);
        return Math.Sqrt(freedom * (y / x));
    }

    // -----------------------------------------------------------------------------------------
    // Fisher transformation and the error function

    // Excel evaluates the textbook formulas literally; FISHER(1e-10) shows their rounding, so do we.
    private static Operand Fisher(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Number(0);
        if (args.Failed)
            return args.Error;
        if (x <= -1 || x >= 1)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(0.5 * Math.Log((1 + x) / (1 - x)));
    }

    private static Operand FisherInv(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var y = args.Number(0);
        if (args.Failed)
            return args.Error;
        var e = Math.Exp(2 * y);
        return CellValue.Number(double.IsInfinity(e) ? 1 : (e - 1) / (e + 1));
    }

    // ERF(lower) or ERF(lower, upper) = erf(upper) - erf(lower).
    private static Operand Erf(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var lower = args.StrictNumber(0);
        var upper = call.Count > 1 ? args.StrictNumber(1) : double.NaN;
        if (args.Failed)
            return args.Error;
        return CellValue.Number(call.Count > 1 ? Special.Erf(upper) - Special.Erf(lower) : Special.Erf(lower));
    }

    private static Operand ErrorFunction(FunctionCall call, bool complement)
    {
        var args = new DistributionArguments(call);
        var x = args.StrictNumber(0);
        if (args.Failed)
            return args.Error;
        return CellValue.Number(complement ? Special.Erfc(x) : Special.Erf(x));
    }
}
