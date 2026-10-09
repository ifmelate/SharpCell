using System;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// Discrete distributions (binomial, negative binomial, hypergeometric, Poisson) and the
/// permutation counts. Counts are truncated toward zero, as Excel does.
/// </summary>
internal static class DistributionDiscreteFunctions
{
    private static readonly ArgumentKind[] Scalars = [ArgumentKind.Value];

    // Beyond 2^53 the trials are no longer whole numbers in a double.
    private const double MaxTrials = 9007199254740992;

    // Below this many terms a range of binomial probabilities is summed term by term.
    private const int MaxSummedTerms = 1000;

    public static void Register(FunctionRegistry registry)
    {
        Add(registry, "BINOM.DIST", 4, 4, BinomDist);
        Add(registry, "BINOMDIST", 4, 4, BinomDist);
        Add(registry, "BINOM.DIST.RANGE", 3, 4, BinomDistRange);
        Add(registry, "BINOM.INV", 3, 3, BinomInv);
        Add(registry, "CRITBINOM", 3, 3, BinomInv);
        Add(registry, "NEGBINOM.DIST", 4, 4, call => NegBinomDist(call, legacy: false));
        Add(registry, "NEGBINOMDIST", 3, 3, call => NegBinomDist(call, legacy: true));
        Add(registry, "HYPGEOM.DIST", 5, 5, call => HypGeomDist(call, legacy: false));
        Add(registry, "HYPGEOMDIST", 4, 4, call => HypGeomDist(call, legacy: true));
        Add(registry, "POISSON.DIST", 3, 3, PoissonDist);
        Add(registry, "POISSON", 3, 3, PoissonDist);
        Add(registry, "PERMUT", 2, 2, Permut);
        Add(registry, "PERMUTATIONA", 2, 2, PermutationA);
    }

    private static void Add(FunctionRegistry registry, string name, int min, int max, FunctionBody body) =>
        registry.Add(new FunctionInfo(name, min, max, Scalars, body));

    // -----------------------------------------------------------------------------------------
    // Binomial

    private static Operand BinomDist(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var successes = args.Count(0);
        var trials = args.Count(1);
        var p = args.Number(2);
        var cumulative = args.Boolean(3);
        if (args.Failed)
            return args.Error;
        if (trials < 0 || trials > MaxTrials || successes < 0 || successes > trials || p < 0 || p > 1)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(cumulative ? BinomialCdf(successes, trials, p) : Special.BinomialTerm(successes, trials, p, 1 - p));
    }

    /// <summary>P(X &lt;= k) for X ~ Binomial(n, p), 0 &lt;= k &lt;= n.</summary>
    private static double BinomialCdf(double k, double n, double p)
    {
        if (k >= n)
            return 1;
        // P(X <= k) = I_{1-p}(n - k, k + 1) = 1 - I_p(k + 1, n - k).
        return Special.BetaRegularized(k + 1, n - k, p, 1 - p, upper: true);
    }

    /// <summary>P(X &gt; k).</summary>
    private static double BinomialUpper(double k, double n, double p) =>
        k >= n ? 0 : Special.BetaRegularized(k + 1, n - k, p, 1 - p);

    // BINOM.DIST.RANGE(trials, p, s, [s2]): P(s <= X <= s2), or P(X = s) without s2.
    private static Operand BinomDistRange(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var trials = args.Integer(0);
        var p = args.Number(1);
        var first = args.Integer(2);
        var last = call.Count > 3 ? args.Integer(3) : first;
        if (args.Failed)
            return args.Error;
        if (trials < 0 || trials > MaxTrials || p < 0 || p > 1 || first < 0 || first > trials || last < first || last > trials)
            return CellValue.Error(ErrorKind.Num);

        var q = 1 - p;
        if (last - first < MaxSummedTerms)
        {
            var sum = 0.0;
            for (var k = first; k <= last; k++)
                sum += Special.BinomialTerm(k, trials, p, q);
            return CellValue.Number(sum);
        }

        // A wide range: the difference of two tails, taken on the side of the mean where they are small.
        var range = first > trials * p
            ? BinomialUpper(first - 1, trials, p) - BinomialUpper(last, trials, p)
            : BinomialCdf(last, trials, p) - (first == 0 ? 0 : BinomialCdf(first - 1, trials, p));
        return CellValue.Number(Math.Max(range, 0));
    }

    // BINOM.INV(trials, p, alpha): the smallest k with P(X <= k) >= alpha.
    private static Operand BinomInv(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var trials = args.Integer(0);
        var p = args.Number(1);
        var alpha = args.Number(2);
        if (args.Failed)
            return args.Error;
        if (trials < 0 || trials > MaxTrials || p < 0 || p >= 1 || alpha <= 0 || alpha >= 1)
            return CellValue.Error(ErrorKind.Num);
        if (p == 0 || trials == 0)
            return CellValue.Number(0);

        var q = 1 - p;
        var first = Math.Pow(q, trials);
        if (trials <= MaxSummedTerms && first > 0)
        {
            // Small counts: add the probabilities up from zero, as the definition reads.
            var term = first;
            var total = term;
            var k = 0.0;
            while (total < alpha && k < trials)
            {
                term *= (trials - k) / (k + 1) * (p / q);
                k++;
                total += term;
            }

            return CellValue.Number(k);
        }

        // Start from the normal approximation and step to the boundary with the exact distribution.
        var mean = trials * p;
        var guess = Math.Floor(mean + Special.NormalQuantile(alpha) * Math.Sqrt(mean * q));
        var result = Math.Clamp(guess, 0, trials);
        var steps = 0;
        while (result < trials && BinomialCdf(result, trials, p) < alpha)
        {
            result++;
            if (++steps % 64 == 0)
                call.Context.CancellationToken.ThrowIfCancellationRequested();
        }

        while (result > 0 && BinomialCdf(result - 1, trials, p) >= alpha)
        {
            result--;
            if (++steps % 64 == 0)
                call.Context.CancellationToken.ThrowIfCancellationRequested();
        }

        return CellValue.Number(result);
    }

    // NEGBINOM.DIST(failures, successes, p, cumulative): failures before the successes-th success.
    private static Operand NegBinomDist(FunctionCall call, bool legacy)
    {
        var args = new DistributionArguments(call);
        var failures = args.Count(0);
        var successes = args.Count(1);
        var p = args.Number(2);
        var cumulative = !legacy && args.Boolean(3);
        if (args.Failed)
            return args.Error;
        if (failures < 0 || successes < 1 || p <= 0 || p >= 1)
            return CellValue.Error(ErrorKind.Num);

        if (cumulative)
            return CellValue.Number(Special.BetaRegularized(successes, failures + 1, p, 1 - p));
        // C(f + s - 1, f) p^s q^f = s / (s + f) * C(s + f, s) p^s q^f.
        return CellValue.Number(successes / (successes + failures) * Special.BinomialTerm(successes, successes + failures, p, 1 - p));
    }

    // -----------------------------------------------------------------------------------------
    // Hypergeometric

    // HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative).
    private static Operand HypGeomDist(FunctionCall call, bool legacy)
    {
        var args = new DistributionArguments(call);
        var x = args.Count(0);
        var sample = args.Count(1);
        var successes = args.Count(2);
        var population = args.Count(3);
        var cumulative = !legacy && args.Boolean(4);
        if (args.Failed)
            return args.Error;
        if (sample <= 0 || sample > population || successes <= 0 || successes > population
            || x < 0 || x > Math.Min(sample, successes) || x < Math.Max(0, sample - population + successes))
            return CellValue.Error(ErrorKind.Num);

        var failures = population - successes;
        return CellValue.Number(cumulative ? HypergeometricCdf(x, successes, failures, sample) : HypergeometricTerm(x, successes, failures, sample));
    }

    // P(X = x) drawing n from r successes and b failures (R's dhyper, through binomial terms).
    private static double HypergeometricTerm(double x, double r, double b, double n)
    {
        var p = n / (r + b);
        var q = (r + b - n) / (r + b);
        var p1 = Special.BinomialTerm(x, r, p, q);
        var p2 = Special.BinomialTerm(n - x, b, p, q);
        var p3 = Special.BinomialTerm(n, r + b, p, q);
        return p1 * p2 / p3;
    }

    // P(X <= x), summing the terms down from x on the side of the mean where the tail is small
    // (R's phyper).
    private static double HypergeometricCdf(double x, double r, double b, double n)
    {
        var lower = true;
        if (x * (r + b) > n * r)
        {
            // Above the mean: P(X <= x) = 1 - P(X' <= n - x - 1) with the roles of r and b swapped.
            (r, b) = (b, r);
            x = n - x - 1;
            lower = false;
        }

        if (x < 0 || x < n - b)
            return lower ? 0 : 1;
        if (x >= r || x >= n)
            return lower ? 1 : 0;

        var term = HypergeometricTerm(x, r, b, n);
        var sum = 0.0;
        var ratio = 1.0;
        for (var k = x; k > 0 && ratio >= 2.220446049250313e-16 * sum; k--)
        {
            ratio *= k * (b - n + k) / ((n + 1 - k) * (r + 1 - k));
            sum += ratio;
        }

        var tail = term * (1 + sum);
        return lower ? tail : 1 - tail;
    }

    // -----------------------------------------------------------------------------------------
    // Poisson and permutations

    private static Operand PoissonDist(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var x = args.Count(0);
        var mean = args.Number(1);
        var cumulative = args.Boolean(2);
        if (args.Failed)
            return args.Error;
        if (x < 0 || mean < 0)
            return CellValue.Error(ErrorKind.Num);
        // P(X <= x) = Q(x + 1, mean).
        return CellValue.Number(cumulative ? Special.GammaQ(x + 1, mean) : Special.PoissonTerm(x, mean));
    }

    // PERMUT(n, k) = n! / (n - k)!.
    private static Operand Permut(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var n = args.Integer(0);
        var k = args.Integer(1);
        if (args.Failed)
            return args.Error;
        if (n <= 0 || k < 0 || k > n)
            return CellValue.Error(ErrorKind.Num);

        // Every factor is at least 1, and from the top they are at least 2 until near the end, so
        // the product overflows within about a thousand steps if it is going to.
        var product = 1.0;
        for (var i = 0.0; i < k && double.IsFinite(product); i++)
            product *= n - i;
        return CellValue.Number(product);
    }

    private static Operand PermutationA(FunctionCall call)
    {
        var args = new DistributionArguments(call);
        var n = args.Integer(0);
        var k = args.Integer(1);
        if (args.Failed)
            return args.Error;
        if (n < 0 || k < 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(Math.Pow(n, k));
    }
}
