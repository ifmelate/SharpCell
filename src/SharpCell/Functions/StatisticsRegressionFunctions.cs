using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// Two-variable statistics (CORREL, SLOPE, COVARIANCE.S...) and least-squares regression
/// (LINEST, LOGEST, TREND, GROWTH).
/// </summary>
internal static class StatisticsRegressionFunctions
{
    // The data are array parameters: before dynamic arrays, CORREL(A1:A9*2,B1:B9) and
    // TREND(LN(A1:A9)) were already calculated as arrays.
    private static readonly ArgumentKind[] TwoArrays = [ArgumentKind.ArrayContext, ArgumentKind.ArrayContext];
    private static readonly ArgumentKind[] EstimateArguments =
        [ArgumentKind.ArrayContext, ArgumentKind.ArrayContext, ArgumentKind.Value, ArgumentKind.Value];
    private static readonly ArgumentKind[] PredictArguments =
        [ArgumentKind.ArrayContext, ArgumentKind.ArrayContext, ArgumentKind.ArrayContext, ArgumentKind.Value];

    // A variable that keeps less than this share of its size once the variables before it are
    // taken out is collinear with them; it is left out of the fit (coefficient 0), as Excel does.
    private const double CollinearityTolerance = 1e-12;

    private const int CancellationCheckInterval = 4096;

    public static void Register(FunctionRegistry registry)
    {
        void Pair(string name, Func<PairMoments, CellValue> statistic) =>
            registry.Add(new FunctionInfo(name, 2, 2, TwoArrays, call => Paired(call, statistic)));

        Pair("COVARIANCE.P", m => m.Count == 0 ? Div0 : CellValue.Number(m.Sxy / m.Count));
        Pair("COVAR", m => m.Count == 0 ? Div0 : CellValue.Number(m.Sxy / m.Count));
        Pair("COVARIANCE.S", m => m.Count < 2 ? Div0 : CellValue.Number(m.Sxy / (m.Count - 1)));
        Pair("CORREL", Correlation);
        Pair("PEARSON", Correlation);
        Pair("RSQ", m => Correlation(m) is { Kind: CellValueKind.Number } r ? CellValue.Number(r.AsNumber() * r.AsNumber()) : Div0);
        Pair("SLOPE", m => m.Sxx == 0 ? Div0 : CellValue.Number(m.Sxy / m.Sxx));
        Pair("INTERCEPT", m => m.Sxx == 0 ? Div0 : CellValue.Number(m.MeanY - m.Sxy / m.Sxx * m.MeanX));
        Pair("STEYX", m => m.Count < 3 || m.Sxx == 0
            ? Div0
            : CellValue.Number(Math.Sqrt(Math.Max(0, m.Syy - m.Sxy * m.Sxy / m.Sxx) / (m.Count - 2))));

        registry.Add(new FunctionInfo("LINEST", 1, 4, EstimateArguments, call => Estimate(call, exponential: false)));
        registry.Add(new FunctionInfo("LOGEST", 1, 4, EstimateArguments, call => Estimate(call, exponential: true)));
        registry.Add(new FunctionInfo("TREND", 1, 4, PredictArguments, call => Predict(call, exponential: false)));
        registry.Add(new FunctionInfo("GROWTH", 1, 4, PredictArguments, call => Predict(call, exponential: true)));
    }

    private static CellValue Div0 => CellValue.Error(ErrorKind.Div0);

    private static CellValue Num => CellValue.Error(ErrorKind.Num);

    /// <summary>Count, means and centred sums of squares and products of paired numbers.</summary>
    private readonly record struct PairMoments(int Count, double MeanX, double MeanY, double Sxx, double Syy, double Sxy);

    private static CellValue Correlation(PairMoments m) =>
        m.Sxx == 0 || m.Syy == 0 ? Div0 : CellValue.Number(m.Sxy / Math.Sqrt(m.Sxx * m.Syy));

    // Positions pair up in row-major order and count only where both hold numbers. Arrays of
    // different sizes are #N/A; an error anywhere is the result. The first argument is y and the
    // second x, as in SLOPE(known_y's, known_x's); the other statistics are symmetric.
    private static Operand Paired(FunctionCall call, Func<PairMoments, CellValue> statistic)
    {
        var context = call.Context;
        if (!Sequence.TryRead(call[0], context, out var first, out var error))
            return error;
        if (!Sequence.TryRead(call[1], context, out var second, out error))
            return error;

        var ys = new List<double>();
        var xs = new List<double>();
        if (Sequence.Pairs(first, second, ys, xs) is { } failure)
            return failure;
        return statistic(Moments(xs, ys));
    }

    // Two passes: means first, then centred sums. No pairs give zero sums, so every statistic
    // above reports #DIV/0!.
    private static PairMoments Moments(List<double> xs, List<double> ys)
    {
        var n = xs.Count;
        if (n == 0)
            return default;
        double meanX = Samples.Mean(xs), meanY = Samples.Mean(ys);
        double sxx = 0, syy = 0, sxy = 0;
        for (var i = 0; i < n; i++)
        {
            var dx = xs[i] - meanX;
            var dy = ys[i] - meanY;
            sxx += dx * dx;
            syy += dy * dy;
            sxy += dx * dy;
        }

        return new PairMoments(n, meanX, meanY, sxx, syy, sxy);
    }

    // LINEST and LOGEST: a row of coefficients, last variable first, then the intercept; with
    // stats, four more rows: standard errors, then (r², se of y), (F, degrees of freedom),
    // (regression and residual sums of squares), the rest of those rows #N/A. LOGEST fits ln y
    // and returns e^coefficient; its statistics stay on the logarithmic scale, as Excel's do.
    private static Operand Estimate(FunctionCall call, bool exponential)
    {
        if (!TryReadModel(call, exponential, out var model, out var error))
            return error;
        var useConstant = call.Boolean(2, true);
        if (useConstant.IsError)
            return useConstant;
        var withStats = call.Boolean(3, false);
        if (withStats.IsError)
            return withStats;

        var fit = Fit.Compute(model, useConstant.AsBoolean(), call.Context);
        var k = model.Variables;
        var result = new CellValue[withStats.AsBoolean() ? 5 : 1, k + 1];
        for (var j = 0; j < k; j++)
            result[0, k - 1 - j] = Coefficient(fit.Coefficients[j], exponential);
        result[0, k] = Coefficient(fit.Intercept, exponential);
        if (!withStats.AsBoolean())
            return CellValue.Array(result);

        for (var r = 1; r < 5; r++)
        {
            for (var c = 0; c < k + 1; c++)
                result[r, c] = CellValue.Error(ErrorKind.NA);
        }

        for (var j = 0; j < k; j++)
            result[1, k - 1 - j] = fit.StandardErrors[j];
        result[1, k] = fit.InterceptError;
        result[2, 0] = fit.RSquared;
        result[2, 1] = fit.StandardErrorOfY;
        result[3, 0] = fit.F;
        result[3, 1] = CellValue.Number(fit.DegreesOfFreedom);
        result[4, 0] = CellValue.Number(fit.RegressionSumOfSquares);
        result[4, 1] = CellValue.Number(fit.ResidualSumOfSquares);
        return CellValue.Array(result);
    }

    private static CellValue Coefficient(double value, bool exponential) => CellValue.Number(exponential ? Math.Exp(value) : value);

    // TREND and GROWTH: the fitted values at new_x's (by default at known_x's). With one variable
    // the result has the shape of new_x's; with several, one value per observation of new_x's,
    // laid out like known_y's.
    private static Operand Predict(FunctionCall call, bool exponential)
    {
        if (!TryReadModel(call, exponential, out var model, out var error))
            return error;
        var useConstant = call.Boolean(3, true);
        if (useConstant.IsError)
            return useConstant;

        double[,] newX;
        if (call.Has(2))
        {
            if (!TryReadNumbers(call[2], call.Context, out newX, out error))
                return error;
        }
        else
        {
            newX = model.XGrid;
        }

        var k = model.Variables;
        int count;
        bool byRows;
        if (k == 1)
        {
            count = newX.Length;
            byRows = true;
        }
        else if (model.ByRows && newX.GetLength(1) == k)
        {
            count = newX.GetLength(0);
            byRows = true;
        }
        else if (!model.ByRows && newX.GetLength(0) == k)
        {
            count = newX.GetLength(1);
            byRows = false;
        }
        else
        {
            return CellValue.Error(ErrorKind.Ref);
        }

        var fit = Fit.Compute(model, useConstant.AsBoolean(), call.Context);
        var result = k == 1
            ? new CellValue[newX.GetLength(0), newX.GetLength(1)]
            : byRows ? new CellValue[count, 1] : new CellValue[1, count];
        var columns = newX.GetLength(1);
        for (var i = 0; i < count; i++)
        {
            var y = fit.Intercept;
            if (k == 1)
                y += fit.Coefficients[0] * newX[i / columns, i % columns];
            else
            {
                for (var j = 0; j < k; j++)
                    y += fit.Coefficients[j] * (byRows ? newX[i, j] : newX[j, i]);
            }

            var value = CellValue.Number(exponential ? Math.Exp(y) : y);
            if (k == 1)
                result[i / columns, i % columns] = value;
            else if (byRows)
                result[i, 0] = value;
            else
                result[0, i] = value;
        }

        return CellValue.Array(result);
    }

    /// <summary>
    /// The observations of a regression: y (as ln y for the exponential fits) and the x
    /// variables, one row per observation.
    /// </summary>
    private sealed class Model(double[] y, double[,] x, double[,] xGrid, bool byRows)
    {
        public double[] Y { get; } = y;

        public double[,] X { get; } = x;

        /// <summary>known_x's as given (or the default 1, 2, 3... in the shape of known_y's), the default of new_x's.</summary>
        public double[,] XGrid { get; } = xGrid;

        /// <summary>Whether observations run down rows (known_y's a column) rather than across columns.</summary>
        public bool ByRows { get; } = byRows;

        public int Observations => Y.Length;

        public int Variables => X.GetLength(1);
    }

    // known_y's and known_x's must be numbers throughout (#VALUE! otherwise). With known_y's a
    // column each column of known_x's is a variable, with known_y's a row each row; a single
    // variable may have any shape equal to known_y's. Other shapes are #REF!.
    private static bool TryReadModel(FunctionCall call, bool exponential, out Model model, out CellValue error)
    {
        model = null!;
        var context = call.Context;
        if (!TryReadNumbers(call[0], context, out var yGrid, out error))
            return false;
        int rows = yGrid.GetLength(0), columns = yGrid.GetLength(1);
        var n = rows * columns;
        var y = new double[n];
        for (var i = 0; i < n; i++)
        {
            y[i] = yGrid[i / columns, i % columns];
            if (exponential)
            {
                if (y[i] <= 0)
                {
                    error = Num;
                    return false;
                }

                y[i] = Math.Log(y[i]);
            }
        }

        double[,] xGrid;
        if (call.Has(1))
        {
            if (!TryReadNumbers(call[1], context, out xGrid, out error))
                return false;
        }
        else
        {
            xGrid = new double[rows, columns];
            for (var i = 0; i < n; i++)
                xGrid[i / columns, i % columns] = i + 1;
        }

        int xRows = xGrid.GetLength(0), xColumns = xGrid.GetLength(1);
        double[,] x;
        bool byRows;
        if (columns == 1 && xRows == rows)
        {
            x = xGrid;
            byRows = true;
        }
        else if (rows == 1 && xColumns == columns)
        {
            x = new double[n, xRows];
            for (var i = 0; i < n; i++)
            {
                for (var j = 0; j < xRows; j++)
                    x[i, j] = xGrid[j, i];
            }

            byRows = false;
        }
        else if (xRows == rows && xColumns == columns)
        {
            x = new double[n, 1];
            for (var i = 0; i < n; i++)
                x[i, 0] = xGrid[i / columns, i % columns];
            byRows = true;
        }
        else
        {
            error = CellValue.Error(ErrorKind.Ref);
            return false;
        }

        model = new Model(y, x, xGrid, byRows);
        return true;
    }

    // A range, array or scalar of numbers only; the first error inside is the result, anything
    // else that is not a number #VALUE!.
    private static bool TryReadNumbers(Operand operand, EvaluationContext context, out double[,] numbers, out CellValue error)
    {
        numbers = null!;
        error = default;
        var value = Evaluator.ToValue(operand, context);
        if (value.IsError)
        {
            error = value;
            return false;
        }

        var array = value.Kind == CellValueKind.Array ? value.AsArray() : new[,] { { value } };
        int rows = array.GetLength(0), columns = array.GetLength(1);
        numbers = new double[rows, columns];
        for (var r = 0; r < rows; r++)
        {
            if (r % CancellationCheckInterval == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            for (var c = 0; c < columns; c++)
            {
                var element = array[r, c];
                if (element.Kind == CellValueKind.Number)
                {
                    numbers[r, c] = element.AsNumber();
                    continue;
                }

                error = element.IsError ? element : CellValue.Error(ErrorKind.Value);
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// A least-squares fit by Householder QR of the (centred, with a constant) observations:
    /// accurate where the normal equations lose digits. Collinear variables are left out.
    /// </summary>
    private sealed class Fit
    {
        public double[] Coefficients { get; private init; } = [];

        public double Intercept { get; private init; }

        public CellValue[] StandardErrors { get; private init; } = [];

        public CellValue InterceptError { get; private init; }

        public CellValue RSquared { get; private init; }

        public CellValue StandardErrorOfY { get; private init; }

        public CellValue F { get; private init; }

        public double DegreesOfFreedom { get; private init; }

        public double RegressionSumOfSquares { get; private init; }

        public double ResidualSumOfSquares { get; private init; }

        public static Fit Compute(Model model, bool useConstant, EvaluationContext context)
        {
            var n = model.Observations;
            var k = model.Variables;

            // The variables with y as one more column, so that each reflection transforms y too.
            var a = new double[n, k + 1];
            for (var i = 0; i < n; i++)
            {
                for (var j = 0; j < k; j++)
                    a[i, j] = model.X[i, j];
                a[i, k] = model.Y[i];
            }

            // means[k] is the mean of y.
            var means = new double[k + 1];
            if (useConstant)
            {
                for (var j = 0; j <= k; j++)
                {
                    for (var i = 0; i < n; i++)
                        means[j] += a[i, j];
                    means[j] /= n;
                    for (var i = 0; i < n; i++)
                        a[i, j] -= means[j];
                }
            }

            var totalSumOfSquares = 0.0;
            for (var i = 0; i < n; i++)
                totalSumOfSquares += a[i, k] * a[i, k];

            // Householder reflections, one per kept variable; kept[r] is the variable of row r of R.
            var kept = new List<int>();
            for (var j = 0; j < k; j++)
            {
                context.CancellationToken.ThrowIfCancellationRequested();
                var r = kept.Count;
                if (r >= n)
                    break;

                var original = 0.0;
                for (var i = 0; i < n; i++)
                    original += model.X[i, j] * model.X[i, j];
                var norm = 0.0;
                for (var i = r; i < n; i++)
                    norm += a[i, j] * a[i, j];
                norm = Math.Sqrt(norm);
                if (norm == 0 || norm <= CollinearityTolerance * Math.Sqrt(original))
                    continue;

                var alpha = a[r, j] > 0 ? -norm : norm;
                var v = new double[n - r];
                for (var i = r; i < n; i++)
                    v[i - r] = a[i, j];
                v[0] -= alpha;
                var vv = 0.0;
                foreach (var e in v)
                    vv += e * e;

                for (var c = j + 1; c <= k; c++)
                {
                    var dot = 0.0;
                    for (var i = r; i < n; i++)
                        dot += v[i - r] * a[i, c];
                    var factor = 2 * dot / vv;
                    for (var i = r; i < n; i++)
                        a[i, c] -= factor * v[i - r];
                }

                a[r, j] = alpha;
                for (var i = r + 1; i < n; i++)
                    a[i, j] = 0;
                kept.Add(j);
            }

            // Back substitution in R b = Q'y over the kept variables.
            var rank = kept.Count;
            var beta = new double[rank];
            for (var row = rank - 1; row >= 0; row--)
            {
                var sum = a[row, k];
                for (var c = row + 1; c < rank; c++)
                    sum -= a[row, kept[c]] * beta[c];
                beta[row] = sum / a[row, kept[row]];
            }

            var coefficients = new double[k];
            for (var i = 0; i < rank; i++)
                coefficients[kept[i]] = beta[i];
            var intercept = 0.0;
            if (useConstant)
            {
                intercept = means[k];
                for (var j = 0; j < k; j++)
                    intercept -= coefficients[j] * means[j];
            }

            var regression = 0.0;
            for (var i = 0; i < rank; i++)
                regression += a[i, k] * a[i, k];
            var residual = 0.0;
            for (var i = rank; i < n; i++)
                residual += a[i, k] * a[i, k];

            var degrees = Math.Max(0, n - rank - (useConstant ? 1 : 0));

            // An exact fit with no degrees of freedom left has zero errors, as in Excel.
            var variance = degrees > 0 ? residual / degrees : 0;
            var inverse = InvertUpper(a, kept);
            var errors = new CellValue[k];
            for (var j = 0; j < k; j++)
                errors[j] = CellValue.Number(0);
            for (var i = 0; i < rank; i++)
            {
                var sum = 0.0;
                for (var c = i; c < rank; c++)
                    sum += inverse[i, c] * inverse[i, c];
                errors[kept[i]] = CellValue.Number(Math.Sqrt(variance * sum));
            }

            var interceptError = CellValue.Error(ErrorKind.NA);
            if (useConstant)
            {
                // Var(b) = s² (1/n + x̄' (X'X)⁻¹ x̄), with (X'X)⁻¹ = R⁻¹R⁻ᵀ.
                var spread = 0.0;
                for (var c = 0; c < rank; c++)
                {
                    var w = 0.0;
                    for (var i = 0; i <= c; i++)
                        w += inverse[i, c] * means[kept[i]];
                    spread += w * w;
                }

                interceptError = CellValue.Number(Math.Sqrt(variance * (1.0 / n + spread)));
            }

            return new Fit
            {
                Coefficients = coefficients,
                Intercept = intercept,
                StandardErrors = errors,
                InterceptError = interceptError,
                RSquared = totalSumOfSquares == 0 ? Num : CellValue.Number(regression / totalSumOfSquares),
                StandardErrorOfY = CellValue.Number(Math.Sqrt(variance)),
                F = rank == 0 || degrees <= 0 || residual == 0 ? Num : CellValue.Number(regression / rank / (residual / degrees)),
                DegreesOfFreedom = degrees,
                RegressionSumOfSquares = regression,
                ResidualSumOfSquares = residual,
            };
        }

        // The inverse of the upper triangular R (rows 0..rank-1 of a, kept columns).
        private static double[,] InvertUpper(double[,] a, List<int> kept)
        {
            var rank = kept.Count;
            var inverse = new double[rank, rank];
            for (var c = 0; c < rank; c++)
            {
                inverse[c, c] = 1 / a[c, kept[c]];
                for (var row = c - 1; row >= 0; row--)
                {
                    var sum = 0.0;
                    for (var m = row + 1; m <= c; m++)
                        sum += a[row, kept[m]] * inverse[m, c];
                    inverse[row, c] = -sum / a[row, kept[row]];
                }
            }

            return inverse;
        }
    }
}
