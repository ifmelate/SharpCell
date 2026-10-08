using SharpCell;

namespace SharpCell.Tests.Evaluation;

public class OperatorTests
{
    private static CellValue Eval(string formula) => new Workbook().Evaluate(formula);

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    public static TheoryData<string, CellValue> Arithmetic => new()
    {
        { "=1+2*3", 7 },
        { "=2^3^2", 64 },
        { "=-2^2", 4 },
        { "=5%", 0.05 },
        { "=10/4", 2.5 },
        { "=7-10", -3 },
        { "=1/0", Err(ErrorKind.Div0) },
        { "=0^0", Err(ErrorKind.Num) },
        { "=0^-1", Err(ErrorKind.Div0) },
        { "=(-8)^(1/3)", Err(ErrorKind.Num) },
        { "=1E308*10", Err(ErrorKind.Num) },
        { "=\"3\"+1", 4 },
        { "=\" 2.5 \"*2", 5 },
        { "=\"a\"+1", Err(ErrorKind.Value) },
        { "=TRUE+1", 2 },
        { "=#N/A+1", Err(ErrorKind.NA) },
        { "=1+#N/A", Err(ErrorKind.NA) },
        { "=\"a\"+#N/A", Err(ErrorKind.Value) },
        { "=#DIV/0!+#N/A", Err(ErrorKind.Div0) },
        { "=-\"2\"", -2 },
        { "=+\"a\"", "a" },
        { "=--TRUE", 1 },
    };

    [Theory]
    [MemberData(nameof(Arithmetic))]
    public void Arithmetic_follows_Excel(string formula, CellValue expected) => Assert.Equal(expected, Eval(formula));

    public static TheoryData<string, CellValue> Concatenation => new()
    {
        { "=\"a\"&1&TRUE", "a1TRUE" },
        { "=1/3&\"\"", "0.333333333333333" },
        { "=\"x\"&#DIV/0!", Err(ErrorKind.Div0) },
        { "=1.5&\"\"", "1.5" },
    };

    [Theory]
    [MemberData(nameof(Concatenation))]
    public void Concatenation_converts_to_text(string formula, CellValue expected) => Assert.Equal(expected, Eval(formula));

    public static TheoryData<string, bool> Comparisons => new()
    {
        { "=1=1", true },
        { "=\"a\"=\"A\"", true },
        { "=\"a\"<\"B\"", true },
        { "=1<\"a\"", true },
        { "=\"a\"<TRUE", true },
        { "=TRUE>1", true },
        { "=FALSE<TRUE", true },
        { "=0.1+0.2=0.3", true },
        { "=\"1\"=1", false },
        { "=2<>2", false },
        { "=3>=3", true },
        { "=3<=2", false },
    };

    [Theory]
    [MemberData(nameof(Comparisons))]
    public void Comparisons_order_types_like_Excel(string formula, bool expected) =>
        Assert.Equal(CellValue.Boolean(expected), Eval(formula));

    [Fact]
    public void Comparison_propagates_the_left_error_first()
    {
        Assert.Equal(Err(ErrorKind.NA), Eval("=#N/A=#DIV/0!"));
    }

    [Fact]
    public void Last_addition_or_subtraction_near_zero_becomes_zero()
    {
        Assert.Equal(CellValue.Number(0), Eval("=0.3-0.2-0.1"));
        Assert.Equal(CellValue.Number((0.3 - 0.2 - 0.1) * 1), Eval("=1*(0.3-0.2-0.1)"));
        Assert.NotEqual(CellValue.Number(0), Eval("=1*(0.3-0.2-0.1)"));
    }

    [Fact]
    public void Arrays_combine_element_wise_with_broadcasting()
    {
        Assert.Equal(CellValue.Array(new CellValue[,] { { 11, 12, 13 }, { 21, 22, 23 } }), Eval("={1,2,3}+{10;20}"));
        Assert.Equal(CellValue.Array(new CellValue[,] { { 2, 4, Err(ErrorKind.NA) } }), Eval("={1,2}+{1,2,3}"));
        Assert.Equal(CellValue.Array(new CellValue[,] { { -1, -2 } }), Eval("=-{1,2}"));
        Assert.Equal(CellValue.Array(new CellValue[,] { { true, false } }), Eval("={1,2}=1"));
    }

    [Fact]
    public void Empty_formula_result_is_zero()
    {
        var wb = new Workbook();
        wb.AddSheet("S");
        Assert.Equal(CellValue.Number(0), wb.Evaluate("=A1"));
    }
}
