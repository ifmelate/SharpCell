using SharpCell;

namespace SharpCell.Tests.Functions;

public class StatisticsRegressionFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public StatisticsRegressionFunctionTests()
    {
        _s = _wb.AddSheet("S");

        // A1:E11: Excel's LINEST documentation example (office buildings): floor space, offices,
        // entrances, age, value.
        double[,] buildings =
        {
            { 2310, 2, 2, 20, 142000 },
            { 2333, 2, 2, 12, 144000 },
            { 2356, 3, 1.5, 33, 151000 },
            { 2379, 3, 2, 43, 150000 },
            { 2402, 2, 3, 53, 139000 },
            { 2425, 4, 2, 23, 169000 },
            { 2448, 2, 1.5, 99, 126000 },
            { 2471, 2, 2, 34, 142900 },
            { 2494, 3, 3, 23, 163000 },
            { 2517, 4, 4, 55, 169000 },
            { 2540, 2, 3, 22, 149000 },
        };
        for (var r = 0; r < buildings.GetLength(0); r++)
        {
            for (var c = 0; c < buildings.GetLength(1); c++)
                _s[r + 1, c + 1].Value = buildings[r, c];
        }

        // G1:H5: pairs with gaps: (1,3), (2,text), (3,7), (empty,5), (5,11).
        _s["G1"].Value = 1;
        _s["H1"].Value = 3;
        _s["G2"].Value = 2;
        _s["H2"].Value = "text";
        _s["G3"].Value = 3;
        _s["H3"].Value = 7;
        _s["H4"].Value = 5;
        _s["G5"].Value = 5;
        _s["H5"].Value = 11;
        _s["J1"].Formula = "=1/0";
        _s["K1"].Value = "y";
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    private static void AssertClose(double expected, CellValue value, double tolerance = 1e-9)
    {
        Assert.True(value.Kind == CellValueKind.Number, $"expected {expected}, got {value}");
        Assert.True(Math.Abs(value.AsNumber() - expected) <= tolerance * Math.Max(1, Math.Abs(expected)),
            $"expected {expected}, got {value.AsNumber()}");
    }

    [Theory]
    [InlineData("=CORREL({3,2,4,5,6},{9,7,12,15,17})", 0.997054485501581)]
    [InlineData("=PEARSON({3,2,4,5,6},{9,7,12,15,17})", 0.997054485501581)]
    [InlineData("=RSQ({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 0.057950191570881)]
    [InlineData("=SLOPE({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 0.305555555555556)]
    [InlineData("=INTERCEPT({2,3,9,1,8},{6,5,11,7,5})", 0.048387096774194)]
    [InlineData("=STEYX({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 3.305718950210041)]
    [InlineData("=COVARIANCE.S({2,4,8},{5,11,12})", 9.666666666666666)]
    [InlineData("=COVARIANCE.P({3,2,4,5,6},{9,7,12,15,17})", 5.2)]
    [InlineData("=COVAR({3,2,4,5,6},{9,7,12,15,17})", 5.2)]
    public void Paired_statistics(string formula, double expected) => AssertClose(expected, Eval(formula), 1e-12);

    [Fact]
    public void Only_positions_with_two_numbers_pair_up()
    {
        // Pairs (1,3), (3,7), (5,11): y = 2x + 1 exactly.
        AssertClose(2, Eval("=SLOPE(H1:H5,G1:G5)"));
        AssertClose(1, Eval("=INTERCEPT(H1:H5,G1:G5)"));
        AssertClose(1, Eval("=CORREL(G1:G5,H1:H5)"));
        AssertClose(16.0 / 3, Eval("=COVARIANCE.P(G1:G5,H1:H5)"));
        AssertClose(16.0 / 3, Eval("=COVARIANCE.P(G:G,H:H)"));
    }

    [Theory]
    [InlineData("=CORREL({1,2,3},{1,2})", ErrorKind.NA)]
    [InlineData("=SLOPE(G1:G5,H1:H4)", ErrorKind.NA)]
    [InlineData("=CORREL({1,2,3},{4,4,4})", ErrorKind.Div0)]
    [InlineData("=SLOPE({1,2,3},{4,4,4})", ErrorKind.Div0)]
    [InlineData("=COVARIANCE.S({5},{10})", ErrorKind.Div0)]
    [InlineData("=COVARIANCE.P({\"a\",\"b\"},{1,2})", ErrorKind.Div0)]
    [InlineData("=STEYX({1,2},{3,4})", ErrorKind.Div0)]
    [InlineData("=RSQ({1},{2})", ErrorKind.Div0)]
    [InlineData("=CORREL(J1:K1,J1:K1)", ErrorKind.Div0)]
    public void Paired_errors(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }

    [Fact]
    public void Paired_covariance_of_one_pair_is_zero()
    {
        Assert.Equal(CellValue.Number(0), Eval("=COVARIANCE.P({5},{10})"));
    }

    [Fact]
    public void Linest_of_one_variable_gives_slope_and_intercept()
    {
        var result = Eval("=LINEST({1,9,5,7},{0,4,2,3})").AsArray();
        Assert.Equal(1, result.GetLength(0));
        Assert.Equal(2, result.GetLength(1));
        AssertClose(2, result[0, 0]);
        AssertClose(1, result[0, 1]);
    }

    [Fact]
    public void Linest_with_several_variables_and_statistics_matches_Excel()
    {
        var result = Eval("=LINEST(E1:E11,A1:D11,TRUE,TRUE)").AsArray();
        Assert.Equal(5, result.GetLength(0));
        Assert.Equal(5, result.GetLength(1));
        double[] coefficients = [-234.2371645, 2553.21066, 12529.76817, 27.64138737, 52317.83051];
        double[] errors = [13.26801148, 530.6691519, 400.0668382, 5.429374042, 12237.3616];
        for (var c = 0; c < 5; c++)
        {
            AssertClose(coefficients[c], result[0, c], 1e-9);
            AssertClose(errors[c], result[1, c], 1e-9);
        }

        AssertClose(0.996747993, result[2, 0], 1e-9);
        AssertClose(970.5784629, result[2, 1], 1e-9);
        AssertClose(459.7536742, result[3, 0], 1e-9);
        AssertClose(6, result[3, 1]);
        AssertClose(1732393319, result[4, 0], 1e-9);
        AssertClose(5652135.316, result[4, 1], 1e-9);
        Assert.Equal(CellValue.Error(ErrorKind.NA), result[2, 2]);
        Assert.Equal(CellValue.Error(ErrorKind.NA), result[4, 4]);
    }

    [Fact]
    public void Linest_without_constant_reports_no_intercept_error()
    {
        var result = Eval("=LINEST({2,4,6.5},{1,2,3},FALSE,TRUE)").AsArray();
        AssertClose(29.5 / 14, result[0, 0]);
        Assert.Equal(CellValue.Number(0), result[0, 1]);
        Assert.Equal(CellValue.Error(ErrorKind.NA), result[1, 1]);
        AssertClose(2, result[3, 1]);
    }

    [Fact]
    public void Linest_drops_a_collinear_variable()
    {
        _s["M1"].Formula = "=LINEST({3;5;7;9;12},{1,2;2,4;3,6;4,8;5,10})";
        _wb.Recalculate();
        var m2 = _s["M1"].Value.AsNumber();
        var m1 = _s["N1"].Value.AsNumber();
        Assert.True(m1 == 0 || m2 == 0, $"one coefficient should be 0: {m1}, {m2}");
        AssertClose(2.2, CellValue.Number(m1 + 2 * m2));
        AssertClose(0.6, _s["O1"].Value);
    }

    [Fact]
    public void Logest_fits_an_exponential_curve()
    {
        var result = Eval("=LOGEST({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16})").AsArray();
        AssertClose(1.463275628, result[0, 0], 1e-9);
        AssertClose(495.3047702, result[0, 1], 1e-8);
        var noConstant = Eval("=LOGEST({1,2,4},{1,2,3},FALSE)").AsArray();
        Assert.Equal(CellValue.Number(1), noConstant[0, 1]);
    }

    [Theory]
    [InlineData("=LINEST({1,\"x\",3})", ErrorKind.Value)]
    [InlineData("=LINEST(G1:G5)", ErrorKind.Value)]
    [InlineData("=LINEST(J1)", ErrorKind.Div0)]
    [InlineData("=LINEST({1,2,3},{1,2})", ErrorKind.Ref)]
    [InlineData("=LOGEST({1,0,3})", ErrorKind.Num)]
    [InlineData("=TREND({1,2,3},{1,2})", ErrorKind.Ref)]
    [InlineData("=GROWTH({1,-2,3})", ErrorKind.Num)]
    [InlineData("=TREND({1;2;3},{1,2;2,3;3,5},{1,2,3})", ErrorKind.Ref)]
    public void Regression_errors(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }

    [Fact]
    public void Trend_predicts_in_the_shape_of_new_x()
    {
        var result = Eval("=TREND({1,2,3},{1,2,3},{4;5})").AsArray();
        Assert.Equal(2, result.GetLength(0));
        AssertClose(4, result[0, 0]);
        AssertClose(5, result[1, 0]);

        var fitted = Eval("=TREND({3,5,8})").AsArray();
        Assert.Equal(3, fitted.GetLength(1));
        AssertClose(17.0 / 6, fitted[0, 0]);
        AssertClose(16.0 / 3, fitted[0, 1]);
    }

    [Fact]
    public void Trend_with_several_variables_gives_one_value_per_observation()
    {
        var result = Eval("=TREND(E1:E11,A1:D11,{2500,3,2,25})").AsArray();
        Assert.Equal((1, 1), (result.GetLength(0), result.GetLength(1)));
        AssertClose(52317.83051 + 27.64138737 * 2500 + 12529.76817 * 3 + 2553.21066 * 2 - 234.2371645 * 25, result[0, 0], 1e-8);
        Assert.Equal(11, Eval("=TREND(E1:E11,A1:D11)").AsArray().GetLength(0));
    }

    [Fact]
    public void Growth_predicts_on_the_exponential_curve()
    {
        var result = Eval("=GROWTH({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16},{17,18})").AsArray();
        AssertClose(320196.7184, result[0, 0], 1e-9);
        AssertClose(468536.054184048, result[0, 1], 1e-9);
    }

    [Fact]
    public void Regression_arrays_are_calculated_as_arrays_in_legacy_formulas()
    {
        _s["M6"].SetFormula("=SLOPE(E1:E11*2,A1:A11)", legacy: true);
        _wb.Recalculate();
        Assert.Equal(_wb.Evaluate("=SLOPE(E1:E11*2,A1:A11)"), _s["M6"].Value);
        Assert.Equal(CellValueKind.Number, _s["M6"].Value.Kind);
    }
}
