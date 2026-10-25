using SharpCell;

namespace SharpCell.Tests.Evaluation;

/// <summary>
/// How COUNTIFS, SUMIFS and the other conditional functions walk their ranges: ranges with gaps,
/// at different offsets, several columns wide, and the order in which values are added.
/// </summary>
public class ConditionalVisitTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public ConditionalVisitTests()
    {
        _s = _wb.AddSheet("S");

        // A1:A8 with gaps; C3:C10 two rows lower, with other gaps.
        _s["A1"].Value = "x";
        _s["A3"].Value = "x";
        _s["A4"].Value = "y";
        _s["A7"].Value = "x";
        _s["C3"].Value = 1;
        _s["C4"].Value = 2;
        _s["C5"].Value = 4;
        _s["C6"].Value = 8;
        _s["C8"].Value = 16;
        _s["C10"].Value = 32;

        // E1:F3 criteria, H2:I4 values: two columns, three rows, one row lower.
        _s["E1"].Value = "k";
        _s["F1"].Value = "k";
        _s["E2"].Value = "k";
        _s["F2"].Value = "k";
        _s["F3"].Value = "k";
        _s["H2"].Value = 1e16;
        _s["I2"].Value = 1;
        _s["H3"].Value = -1e16;
        _s["I3"].Value = 1;
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    [Theory]
    [InlineData("=SUMIFS(C3:C10,A1:A8,\"x\")", 5)]
    [InlineData("=SUMIFS(C3:C10,A1:A8,\"\")", 50)]
    [InlineData("=SUMIF(A1:A8,\"\",C3)", 50)]
    [InlineData("=SUMIFS(C3:C10,A1:A8,\"<>\")", 13)]
    [InlineData("=COUNTIFS(A1:A8,\"\",C3:C10,\"\")", 1)]
    [InlineData("=COUNTIFS(A1:A8,\"<>\",C3:C10,\"\")", 1)]
    [InlineData("=COUNTIFS(A1:A8,\"\",C3:C10,\">1\")", 3)]
    [InlineData("=COUNTIFS(A1:A8,\"x\",A1:A8,\"<>\")", 3)]
    [InlineData("=COUNTIFS(A:A,\"\",C:C,\"\")", 1048568)]
    [InlineData("=COUNTIFS(A:A,\"\",C:C,\"<>\")", 4)]
    [InlineData("=AVERAGEIFS(C3:C10,A1:A8,\"\")", 50.0 / 3)]
    [InlineData("=MAXIFS(C3:C10,A1:A8,\"\")", 32)]
    [InlineData("=MINIFS(C3:C10,A1:A8,\"x\")", 1)]
    [InlineData("=AVERAGEIF(A1:A8,\"\",C3)", 50.0 / 3)]
    public void Ranges_with_gaps_at_different_offsets(string formula, double expected)
    {
        var value = Eval(formula);
        Assert.Equal(CellValueKind.Number, value.Kind);
        Assert.Equal(expected, value.AsNumber(), 12);
    }

    [Theory]
    [InlineData("=COUNTIFS(E1:F3,\"k\",H2:I4,\"<>\")", 4)]
    [InlineData("=COUNTIFS(E1:F3,\"\",H2:I4,\"\")", 1)]
    [InlineData("=COUNTIFS(E1:F3,\"k\")", 5)]
    [InlineData("=COUNTIFS(H2:I4,\"\")", 2)]
    [InlineData("=MAXIFS(H2:I4,E1:F3,\"k\")", 1e16)]
    [InlineData("=MINIFS(H2:I4,E1:F3,\"k\")", -1e16)]
    public void Ranges_of_several_rows_and_columns(string formula, double expected)
    {
        Assert.Equal(CellValue.Number(expected), Eval(formula));
    }

    // Row by row: 1e16 + 1 loses the 1, then -1e16 and 1 leave 1. Column by column would give 2.
    [Fact]
    public void Values_are_added_row_by_row()
    {
        Assert.Equal(CellValue.Number(1), Eval("=SUMIFS(H2:I4,E1:F3,\"k\")"));
        Assert.Equal(CellValue.Number(1), Eval("=SUMIF(E1:F3,\"k\",H2)"));
        Assert.Equal(CellValue.Number(0.25), Eval("=AVERAGEIFS(H2:I4,E1:F3,\"k\")"));
    }

    // Ranges much larger than the sheet are walked by their stored cells only; the order holds there too.
    [Fact]
    public void Values_of_ranges_down_to_the_last_row_are_added_row_by_row()
    {
        Assert.Equal(CellValue.Number(1), Eval("=SUMIFS(H2:I1048576,E1:F1048575,\"k\")"));
        Assert.Equal(CellValue.Number(1), Eval("=SUMIFS(H2:I1048576,E1:F1048575,\"<>\",H2:I1048576,\"<>\")"));
        Assert.Equal(CellValue.Number(4), Eval("=COUNTIFS(E1:F1048575,\"k\",H2:I1048576,\"<>\")"));
        Assert.Equal(CellValue.Number(2097145), Eval("=COUNTIFS(E1:F1048575,\"\",H2:I1048576,\"\")"));
    }

    [Theory]
    [InlineData("=SUMIFS(H2:I4,E1:E3,\"k\")")]
    [InlineData("=COUNTIFS(E1:F3,\"k\",H2:H4,\"<>\")")]
    [InlineData("=COUNTIFS(E1:F3,\"k\",H2:I3,\"<>\")")]
    [InlineData("=MAXIFS(H2:I3,E1:F3,\"k\")")]
    public void Shapes_that_differ_in_rows_or_columns_are_VALUE(string formula)
    {
        Assert.Equal(CellValue.Error(ErrorKind.Value), Eval(formula));
    }

    [Fact]
    public void An_error_counts_only_where_the_criteria_hold()
    {
        _s["C7"].Formula = "=1/0";
        Assert.Equal(CellValue.Number(5), Eval("=SUMIFS(C3:C10,A1:A8,\"x\")"));
        Assert.Equal(CellValue.Error(ErrorKind.Div0), Eval("=SUMIFS(C3:C10,A1:A8,\"\")"));
        Assert.Equal(CellValue.Error(ErrorKind.Div0), Eval("=MINIFS(C3:C10,A1:A8,\"\")"));
    }

    [Fact]
    public void Criteria_ranges_of_formulas_are_calculated_first()
    {
        _s["J1"].Value = 1;
        _s["J2"].Value = 2;
        _s["J3"].Value = 3;
        _s["K1"].Formula = "=IF(J1>1,\"x\",\"y\")";
        _s["K2"].Formula = "=IF(J2>1,\"x\",\"y\")";
        _s["K3"].Formula = "=IF(J3>1,\"x\",\"y\")";
        _s["L1"].Formula = "=J1*10";
        _s["L2"].Formula = "=J2*10";
        _s["L3"].Formula = "=J3*10";
        _s["M1"].Formula = "=SUMIFS(L1:L3,K1:K3,\"x\")";
        _s["M2"].Formula = "=COUNTIFS(K1:K3,\"x\",L1:L3,\">15\")";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(50), _s["M1"].Value);
        Assert.Equal(CellValue.Number(2), _s["M2"].Value);

        _s["J1"].Value = 5;
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(100), _s["M1"].Value);
        Assert.Equal(CellValue.Number(3), _s["M2"].Value);
    }

    // Excel takes only references here; the engine also takes arrays, where every position holds a value.
    [Fact]
    public void An_array_range_is_walked_with_a_reference()
    {
        Assert.Equal(CellValue.Number(3), Eval("=COUNTIFS({\"x\";\"\";\"x\";\"x\"},\"x\",A1:A4,\"<>\")"));
        Assert.Equal(CellValue.Number(1), Eval("=COUNTIFS({\"x\";\"\";\"x\";\"x\"},\"\",A1:A4,\"\")"));
        Assert.Equal(CellValue.Number(14), Eval("=SUMIFS(C3:C6,{1;2;3;4},\">1\")"));
    }
}
