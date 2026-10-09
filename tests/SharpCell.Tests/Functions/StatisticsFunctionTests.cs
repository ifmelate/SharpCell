using SharpCell;

namespace SharpCell.Tests.Functions;

public class StatisticsFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public StatisticsFunctionTests()
    {
        _s = _wb.AddSheet("S");

        // A1:A6 mixes kinds: 1, 2, TRUE, "text", "4", 5 (A7 empty).
        _s["A1"].Value = 1;
        _s["A2"].Value = 2;
        _s["A3"].Value = true;
        _s["A4"].Value = "text";
        _s["A5"].Value = "4";
        _s["A6"].Value = 5;

        // B1:B10: the sample from Excel's STDEV documentation.
        double[] sample = [1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299];
        for (var i = 0; i < sample.Length; i++)
            _s[$"B{i + 1}"].Value = sample[i];

        _s["C1"].Value = 1;
        _s["C2"].Formula = "=1/0";
        _s["C3"].Formula = "=\"\"";
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    private void AssertNumber(double expected, string formula, int digits = 9)
    {
        var value = Eval(formula);
        Assert.True(value.Kind == CellValueKind.Number, $"{formula} gave {value}");
        Assert.Equal(expected, value.AsNumber(), digits);
    }

    [Theory]
    [InlineData("=STDEV.S(B1:B10)", 27.463915719843)]
    [InlineData("=STDEV(B1:B10)", 27.463915719843)]
    [InlineData("=STDEV.P(B1:B10)", 26.054558142482)]
    [InlineData("=STDEVP(B1:B10)", 26.054558142482)]
    [InlineData("=VAR.S(B1:B10)", 754.266666666667)]
    [InlineData("=VAR(B1:B10)", 754.266666666667)]
    [InlineData("=VAR.P(B1:B10)", 678.84)]
    [InlineData("=VARP(B1:B10)", 678.84)]
    [InlineData("=VAR.P({1,2,3})", 0.666666666666667)]
    [InlineData("=VAR.S(1,{1,2,3},5)", 2.8)]
    public void Variances_like_Excel(string formula, double expected) => AssertNumber(expected, formula);

    [Fact]
    public void Plain_variance_skips_text_and_logicals_in_ranges()
    {
        // Numbers 1, 2, 5.
        AssertNumber(2.888888888888889, "=VAR.P(A1:A6)");
        AssertNumber(4.333333333333333, "=VAR.S(A1:A6)");
    }

    [Fact]
    public void A_variants_count_logicals_as_numbers_and_text_as_zero()
    {
        // Values 1, 2, 1, 0, 0, 5.
        AssertNumber(1.5, "=AVERAGEA(A1:A6)");
        AssertNumber(3.5, "=VARA(A1:A6)");
        AssertNumber(Math.Sqrt(3.5), "=STDEVA(A1:A6)");
        AssertNumber(35.0 / 12, "=VARPA(A1:A6)");
        AssertNumber(Math.Sqrt(35.0 / 12), "=STDEVPA(A1:A6)");
        AssertNumber(0, "=MINA(A3:A5)");
        AssertNumber(1, "=MAXA(A3:A5)");
        AssertNumber(0.5, "=AVERAGEA({1,\"a\"})");
    }

    [Theory]
    [InlineData("=AVERAGEA(1,2,\"5\",TRUE)", 2.25)]
    [InlineData("=MINA(4,\"-1\",TRUE)", -1)]
    [InlineData("=MAXA(A7)", 0)]
    [InlineData("=MINA(A4)", 0)]
    [InlineData("=AVERAGEA(1,)", 0.5)]
    public void A_variants_coerce_direct_arguments(string formula, double expected) => AssertNumber(expected, formula);

    [Theory]
    [InlineData("=AVERAGEA(1,\"x\")", ErrorKind.Value)]
    [InlineData("=AVERAGEA(A7)", ErrorKind.Div0)]
    [InlineData("=AVERAGEA(C1:C2)", ErrorKind.Div0)]
    [InlineData("=MAXA(C1:C2)", ErrorKind.Div0)]
    [InlineData("=VAR.S(1)", ErrorKind.Div0)]
    [InlineData("=STDEV.S(A4:A5)", ErrorKind.Div0)]
    [InlineData("=VAR.P(A4)", ErrorKind.Div0)]
    [InlineData("=VARA(1)", ErrorKind.Div0)]
    [InlineData("=STDEV.P(1,\"x\")", ErrorKind.Value)]
    [InlineData("=STDEV.P(C1:C2)", ErrorKind.Div0)]
    public void Variance_and_average_errors(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }

    [Theory]
    [InlineData("=AVEDEV(4,5,6,7,5,4,3)", 1.020408163265306)]
    [InlineData("=DEVSQ(4,5,8,7,11,4,3)", 48)]
    [InlineData("=GEOMEAN(4,5,8,7,11,4,3)", 5.476986969656962)]
    [InlineData("=HARMEAN(4,5,8,7,11,4,3)", 5.028375962061728)]
    [InlineData("=KURT(3,4,5,2,3,4,5,6,4,7)", -0.151799637208419)]
    [InlineData("=SKEW(3,4,5,2,3,4,5,6,4,7)", 0.359543071407217)]
    [InlineData("=SKEW.P(3,4,5,2,3,4,5,6,4,7)", 0.303193339354144)]
    [InlineData("=MEDIAN(1,2,3,4,5)", 3)]
    [InlineData("=MEDIAN(6,1,2,5,3,4)", 3.5)]
    [InlineData("=MEDIAN(A1:A6)", 2)]
    [InlineData("=AVEDEV(A1:A6)", 14.0 / 9)]
    [InlineData("=GEOMEAN(\"4\",9)", 6)]
    public void Moments_and_means(string formula, double expected) => AssertNumber(expected, formula);

    [Theory]
    [InlineData("=AVEDEV(A4)", ErrorKind.Num)]
    [InlineData("=DEVSQ(A4)", ErrorKind.Num)]
    [InlineData("=GEOMEAN(4,0)", ErrorKind.Num)]
    [InlineData("=GEOMEAN(4,-1)", ErrorKind.Num)]
    [InlineData("=HARMEAN(4,0)", ErrorKind.Num)]
    [InlineData("=HARMEAN(A4)", ErrorKind.Num)]
    [InlineData("=KURT(1,2,3)", ErrorKind.Div0)]
    [InlineData("=KURT(2,2,2,2)", ErrorKind.Div0)]
    [InlineData("=SKEW(1,2)", ErrorKind.Div0)]
    [InlineData("=SKEW(3,3,3)", ErrorKind.Div0)]
    [InlineData("=SKEW.P(3)", ErrorKind.Div0)]
    [InlineData("=MEDIAN(A4)", ErrorKind.Num)]
    [InlineData("=MEDIAN(C1:C2)", ErrorKind.Div0)]
    [InlineData("=GEOMEAN(1,\"x\")", ErrorKind.Value)]
    public void Moment_errors(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }

    [Theory]
    [InlineData("=MODE.SNGL(5.6,4,4,3,2,4)", 4)]
    [InlineData("=MODE(5.6,4,4,3,2,4)", 4)]
    [InlineData("=MODE.SNGL({1,2,3,2,2,3})", 2)]
    [InlineData("=MODE.SNGL(3,1,1,3)", 3)]
    public void Mode_is_the_most_frequent_and_first_seen(string formula, double expected) => AssertNumber(expected, formula);

    [Fact]
    public void Mode_without_repeats_is_NA()
    {
        Assert.Equal(CellValue.Error(ErrorKind.NA), Eval("=MODE.SNGL(1,2,3)"));
        Assert.Equal(CellValue.Error(ErrorKind.NA), Eval("=MODE.MULT({1,2,3,4})"));
        Assert.Equal(CellValue.Error(ErrorKind.NA), Eval("=MODE(A4)"));
    }

    [Fact]
    public void Mode_mult_lists_every_mode_in_order_of_appearance_as_a_column()
    {
        var result = Eval("=MODE.MULT({2,2,1,1,3,3,4})").AsArray();
        Assert.Equal(3, result.GetLength(0));
        Assert.Equal(1, result.GetLength(1));
        Assert.Equal([CellValue.Number(2), CellValue.Number(1), CellValue.Number(3)], new[] { result[0, 0], result[1, 0], result[2, 0] });
    }

    [Theory]
    [InlineData("=COUNTBLANK(A1:A10)", 4)]
    [InlineData("=COUNTBLANK(C1:C4)", 2)]
    [InlineData("=COUNTBLANK(A7)", 1)]
    [InlineData("=COUNTBLANK(A1)", 0)]
    [InlineData("=COUNTBLANK(D:D)", 1048576)]
    [InlineData("=COUNTBLANK(A:A)", 1048570)]
    public void Countblank_counts_empty_cells_and_empty_text(string formula, double expected)
    {
        Assert.Equal(CellValue.Number(expected), Eval(formula));
    }

    [Fact]
    public void Whole_column_statistics_read_only_stored_cells()
    {
        AssertNumber(754.266666666667, "=VAR.S(B:B)");
        AssertNumber(1320, "=MEDIAN(B:B)");
    }
}
