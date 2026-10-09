using SharpCell;

namespace SharpCell.Tests.Functions;

public class StatisticsRankFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public StatisticsRankFunctionTests()
    {
        _s = _wb.AddSheet("S");

        // A1:A10 = 1..10; B1:B8 = 7, 3.5, 3.5, 1, 2, then TRUE, "9", text.
        for (var i = 1; i <= 10; i++)
            _s[$"A{i}"].Value = i;
        double[] b = [7, 3.5, 3.5, 1, 2];
        for (var i = 0; i < b.Length; i++)
            _s[$"B{i + 1}"].Value = b[i];
        _s["B6"].Value = true;
        _s["B7"].Value = "9";
        _s["B8"].Value = "text";
        _s["C1"].Value = 1;
        _s["C2"].Formula = "=1/0";
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    private void AssertNumber(double expected, string formula)
    {
        var value = Eval(formula);
        Assert.True(value.Kind == CellValueKind.Number, $"{formula} gave {value}");
        Assert.Equal(expected, value.AsNumber(), 12);
    }

    private static CellValue[] Column(CellValue value)
    {
        var array = value.AsArray();
        Assert.Equal(1, array.GetLength(1));
        var column = new CellValue[array.GetLength(0)];
        for (var i = 0; i < column.Length; i++)
            column[i] = array[i, 0];
        return column;
    }

    [Theory]
    [InlineData("=LARGE(B1:B8,1)", 7)]
    [InlineData("=LARGE(B1:B8,2)", 3.5)]
    [InlineData("=LARGE(B1:B8,3)", 3.5)]
    [InlineData("=SMALL(B1:B8,1)", 1)]
    [InlineData("=SMALL(B1:B8,5)", 7)]
    [InlineData("=SMALL(B1:B8,1.2)", 2)]
    [InlineData("=LARGE({1,2,3,-8},4)", -8)]
    [InlineData("=SMALL(3.4,1)", 3.4)]
    [InlineData("=SMALL(\"3.4\",1)", 3.4)]
    [InlineData("=SMALL(TRUE,1)", 1)]
    [InlineData("=SMALL({1,\"-1\"},1)", 1)]
    [InlineData("=LARGE(A:A,1)", 10)]
    public void Large_and_small(string formula, double expected) => AssertNumber(expected, formula);

    [Theory]
    [InlineData("=SMALL(B1:B8,0)", ErrorKind.Num)]
    [InlineData("=SMALL(B1:B8,6)", ErrorKind.Num)]
    [InlineData("=LARGE(B6:B8,1)", ErrorKind.Num)]
    [InlineData("=SMALL(\"Hello\",1)", ErrorKind.Value)]
    [InlineData("=LARGE(C1:C2,1)", ErrorKind.Div0)]
    [InlineData("=LARGE(A1:A10,\"x\")", ErrorKind.Value)]
    public void Large_and_small_errors(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }

    [Fact]
    public void An_array_of_positions_gives_an_array()
    {
        var result = Eval("=LARGE(A1:A10,{1;2;3})");
        Assert.Equal([CellValue.Number(10), CellValue.Number(9), CellValue.Number(8)], Column(result));
    }

    [Theory]
    [InlineData("=PERCENTILE.INC(A1:A10,0)", 1)]
    [InlineData("=PERCENTILE.INC(A1:A10,1)", 10)]
    [InlineData("=PERCENTILE.INC(A1:A10,0.4)", 4.6)]
    [InlineData("=PERCENTILE(A1:A10,0.4)", 4.6)]
    [InlineData("=PERCENTILE.EXC(A1:A10,0.4)", 4.4)]
    [InlineData("=PERCENTILE.EXC(A1:A10,0.11)", 1.21)]
    [InlineData("=PERCENTILE.INC({5},0.3)", 5)]
    [InlineData("=QUARTILE.INC(A1:A10,0)", 1)]
    [InlineData("=QUARTILE.INC(A1:A10,1.3)", 3.25)]
    [InlineData("=QUARTILE(A1:A10,2)", 5.5)]
    [InlineData("=QUARTILE.INC(A1:A10,4.5)", 10)]
    [InlineData("=QUARTILE.EXC(A1:A10,1)", 2.75)]
    [InlineData("=QUARTILE.EXC(A1:A10,3)", 8.25)]
    public void Percentiles_and_quartiles(string formula, double expected) => AssertNumber(expected, formula);

    [Theory]
    [InlineData("=PERCENTILE.INC(A1:A10,-0.1)")]
    [InlineData("=PERCENTILE.INC(A1:A10,1.1)")]
    [InlineData("=PERCENTILE.EXC(A1:A10,0)")]
    [InlineData("=PERCENTILE.EXC(A1:A10,1)")]
    [InlineData("=PERCENTILE.EXC(A1:A10,0.05)")]
    [InlineData("=PERCENTILE.EXC(A1:A10,0.95)")]
    [InlineData("=PERCENTILE.INC(B8,0.5)")]
    [InlineData("=QUARTILE.INC(A1:A10,-1)")]
    [InlineData("=QUARTILE.INC(A1:A10,5)")]
    [InlineData("=QUARTILE.EXC(A1:A10,0)")]
    [InlineData("=QUARTILE.EXC(A1:A10,4)")]
    public void Percentiles_out_of_range_are_NUM(string formula)
    {
        Assert.Equal(CellValue.Error(ErrorKind.Num), Eval(formula));
    }

    [Theory]
    [InlineData("=PERCENTRANK.INC(A1:A10,7)", 0.666)]
    [InlineData("=PERCENTRANK(A1:A10,7)", 0.666)]
    [InlineData("=PERCENTRANK.INC(A1:A10,5.12,6)", 0.457777)]
    [InlineData("=PERCENTRANK.INC(A1:A10,1)", 0)]
    [InlineData("=PERCENTRANK.INC(A1:A10,10)", 1)]
    [InlineData("=PERCENTRANK.EXC(A1:A10,9,2)", 0.81)]
    [InlineData("=PERCENTRANK.EXC(A1:A10,2.5,7)", 0.2272727)]
    [InlineData("=PERCENTRANK.INC(B1:B5,3.5)", 0.5)]
    [InlineData("=PERCENTRANK.INC({4},4)", 1)]
    public void Percent_ranks_are_truncated_to_the_significance(string formula, double expected) => AssertNumber(expected, formula);

    [Theory]
    [InlineData("=PERCENTRANK.INC(A1:A10,0.5)", ErrorKind.NA)]
    [InlineData("=PERCENTRANK.EXC(A1:A10,11)", ErrorKind.NA)]
    [InlineData("=PERCENTRANK.INC(A1:A10,5,0)", ErrorKind.Num)]
    [InlineData("=PERCENTRANK.INC(B8,5)", ErrorKind.Num)]
    public void Percent_rank_errors(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }

    [Theory]
    [InlineData("=RANK.EQ(3.5,B1:B8)", 2)]
    [InlineData("=RANK(3.5,B1:B8)", 2)]
    [InlineData("=RANK.EQ(3.5,B1:B8,1)", 3)]
    [InlineData("=RANK.AVG(3.5,B1:B8)", 2.5)]
    [InlineData("=RANK.AVG(3.5,B1:B8,1)", 3.5)]
    [InlineData("=RANK.EQ(7,B1:B8)", 1)]
    [InlineData("=RANK.AVG(1,B1:B8,TRUE)", 1)]
    public void Ranks(string formula, double expected) => AssertNumber(expected, formula);

    [Fact]
    public void Rank_of_a_missing_number_is_NA()
    {
        Assert.Equal(CellValue.Error(ErrorKind.NA), Eval("=RANK.EQ(9,B1:B8)"));
        Assert.Equal(CellValue.Error(ErrorKind.Div0), Eval("=RANK.EQ(1,C1:C2)"));
    }

    [Theory]
    [InlineData("=TRIMMEAN(A1:A10,0)", 5.5)]
    [InlineData("=TRIMMEAN(A1:A10,0.2)", 5.5)]
    [InlineData("=TRIMMEAN({1,2,3,4,100},0.4)", 3)]
    [InlineData("=TRIMMEAN({1,2,3,4,100},0.39)", 22)]
    [InlineData("=TRIMMEAN(B1:B8,0.5)", 3)]
    public void Trimmed_means(string formula, double expected) => AssertNumber(expected, formula);

    [Theory]
    [InlineData("=TRIMMEAN(A1:A10,-0.1)", ErrorKind.Num)]
    [InlineData("=TRIMMEAN(A1:A10,1)", ErrorKind.Num)]
    [InlineData("=TRIMMEAN(B8,0.1)", ErrorKind.Num)]
    [InlineData("=TRIMMEAN(C1:C2,0.1)", ErrorKind.Div0)]
    public void Trimmed_mean_errors(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }

    [Fact]
    public void Frequency_counts_into_bins_in_their_own_order()
    {
        var result = Eval("=FREQUENCY({1,7,23,12,13,90,8,78,24,16},{10,20,15,5})");
        Assert.Equal([CellValue.Number(2), CellValue.Number(1), CellValue.Number(2), CellValue.Number(1), CellValue.Number(4)], Column(result));
    }

    [Fact]
    public void Frequency_gives_duplicate_bins_to_the_first_and_skips_non_numbers()
    {
        var result = Eval("=FREQUENCY(B1:B8,{3.5,3.5,1})");
        Assert.Equal([CellValue.Number(3), CellValue.Number(0), CellValue.Number(1), CellValue.Number(1)], Column(result));
    }

    [Fact]
    public void Frequency_without_bins_counts_everything()
    {
        Assert.Equal([CellValue.Number(5)], Column(Eval("=FREQUENCY(B1:B8,B8)")));
        Assert.Equal(CellValue.Error(ErrorKind.Div0), Eval("=FREQUENCY(C1:C2,{1})"));
    }

    [Fact]
    public void Array_parameters_are_calculated_as_arrays_in_legacy_formulas()
    {
        _s["D1"].SetFormula("=LARGE(A1:A10*2,1)", legacy: true);
        _s["D2"].SetFormula("=SUM(FREQUENCY(A1:A10*2,10))", legacy: true);
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(20), _s["D1"].Value);
        Assert.Equal(CellValue.Number(10), _s["D2"].Value);
    }
}
