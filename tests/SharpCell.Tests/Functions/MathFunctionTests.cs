using System;
using SharpCell;

namespace SharpCell.Tests.Functions;

public class MathFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public MathFunctionTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = 4;
        _s["A2"].Value = "8";
        _s["A3"].Value = true;
        _s["A4"].Value = "text";
        _s["A6"].Value = 12;
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    private void AssertNumber(double expected, string formula, double tolerance = 1e-12)
    {
        var actual = Eval(formula);
        Assert.True(actual.Kind == CellValueKind.Number, $"{formula}: {actual}");
        Assert.True(Math.Abs(actual.AsNumber() - expected) <= tolerance * Math.Max(1, Math.Abs(expected)), $"{formula}: {actual}, expected {expected}");
    }

    private void AssertError(ErrorKind expected, string formula) => Assert.Equal(CellValue.Error(expected), Eval(formula));

    [Theory]
    [InlineData("=PRODUCT(2,3,\"4\")", 24)]
    [InlineData("=PRODUCT(A1:A4)", 4)]
    [InlineData("=PRODUCT(A4:A5)", 0)]
    [InlineData("=SUMSQ(A1:A6,2)", 164)]
    [InlineData("=SIGN(-3)", -1)]
    [InlineData("=SQRT(16)", 4)]
    [InlineData("=SQRTPI(1)", 1.7724538509055159)]
    [InlineData("=EXP(1)", Math.E)]
    [InlineData("=EXP(-745)", 0)]
    [InlineData("=LN(1)", 0)]
    [InlineData("=LOG(8,2)", 3)]
    [InlineData("=LOG(1000)", 3)]
    [InlineData("=LOG10(0.01)", -2)]
    [InlineData("=POWER(-8,1/3)", -2)]
    [InlineData("=MOD(-3,2)", 1)]
    [InlineData("=MOD(3,-2)", -1)]
    [InlineData("=MOD(5,1E-10)", 9.999981783901343E-11)]
    [InlineData("=QUOTIENT(-7,2)", -3)]
    [InlineData("=PI()", Math.PI)]
    [InlineData("=FACT(5.9)", 120)]
    [InlineData("=FACT(0)", 1)]
    [InlineData("=FACTDOUBLE(7)", 105)]
    [InlineData("=FACTDOUBLE(-1)", 1)]
    [InlineData("=COMBIN(5,2)", 10)]
    [InlineData("=COMBIN(1000,200)", 6.617155560659313E+215)]
    [InlineData("=COMBINA(10,10)", 92378)]
    [InlineData("=COMBINA(0,0)", 1)]
    [InlineData("=MULTINOMIAL(2,3,4)", 1260)]
    [InlineData("=MULTINOMIAL({1.7,1.2,3.4,2.3})", 420)]
    [InlineData("=GCD(24,36,\"8\")", 4)]
    [InlineData("=LCM(4,6,10.9)", 60)]
    [InlineData("=GCD(A5:A6)", 12)]
    [InlineData("=LCM(A5:A6)", 0)]
    [InlineData("=LCM(A5,A6)", 12)]
    [InlineData("=SERIESSUM(2,0,1,{1,2,3})", 17)]
    public void Computes_like_Excel(string formula, double expected)
    {
        AssertNumber(expected, formula);
    }

    [Theory]
    [InlineData("=SQRT(-1)", ErrorKind.Num)]
    [InlineData("=LN(0)", ErrorKind.Num)]
    [InlineData("=LOG(10,1)", ErrorKind.Div0)]
    [InlineData("=LOG(-1)", ErrorKind.Num)]
    [InlineData("=POWER(0,0)", ErrorKind.Num)]
    [InlineData("=POWER(0,-1)", ErrorKind.Div0)]
    [InlineData("=EXP(1000)", ErrorKind.Num)]
    [InlineData("=MOD(1,0)", ErrorKind.Div0)]
    [InlineData("=QUOTIENT(1,0)", ErrorKind.Div0)]
    [InlineData("=FACT(-1)", ErrorKind.Num)]
    [InlineData("=FACT(171)", ErrorKind.Num)]
    [InlineData("=COMBIN(3,10)", ErrorKind.Num)]
    [InlineData("=COMBIN(-1,1)", ErrorKind.Num)]
    [InlineData("=COMBIN(2000,1000)", ErrorKind.Num)]
    [InlineData("=COMBINA(0,5)", ErrorKind.Num)]
    [InlineData("=MULTINOMIAL(-1,2)", ErrorKind.Num)]
    [InlineData("=GCD(-4,2)", ErrorKind.Num)]
    [InlineData("=GCD(1E+100)", ErrorKind.Num)]
    [InlineData("=LCM(2^31-1,2^30-1)", ErrorKind.Num)]
    [InlineData("=SERIESSUM(0,0,2,{1,2,3})", ErrorKind.Num)]
    [InlineData("=SERIESSUM(1,0,1,{1,\"a\"})", ErrorKind.Value)]
    [InlineData("=SQRT(\"text\")", ErrorKind.Value)]
    [InlineData("=SIGN(1/0)", ErrorKind.Div0)]
    public void Fails_like_Excel(string formula, ErrorKind expected)
    {
        AssertError(expected, formula);
    }

    [Theory]
    [InlineData("=QUOTIENT(A3,1)")]
    [InlineData("=SQRTPI(TRUE)")]
    [InlineData("=FACTDOUBLE(A3)")]
    [InlineData("=GCD(A3,2)")]
    [InlineData("=LCM({4,TRUE})")]
    [InlineData("=MULTINOMIAL(TRUE,1)")]
    [InlineData("=SERIESSUM(TRUE,0,1,{1})")]
    [InlineData("=RANDBETWEEN(TRUE,2)")]
    public void Analysis_ToolPak_functions_refuse_booleans(string formula)
    {
        AssertError(ErrorKind.Value, formula);
    }

    [Fact]
    public void Other_functions_take_booleans_as_numbers()
    {
        AssertNumber(1, "=MOD(A3,2)");
        AssertNumber(1, "=FACT(TRUE)");
        AssertNumber(1, "=SQRT(A3)");
    }

    [Fact]
    public void An_argument_left_empty_is_NA_for_divisor_functions()
    {
        AssertError(ErrorKind.NA, "=GCD(,4)");
        AssertError(ErrorKind.NA, "=LCM(4,)");
    }

    [Fact]
    public void Divisor_functions_of_only_empty_cells_are_VALUE()
    {
        AssertError(ErrorKind.Value, "=GCD(A5,A7)");
        AssertNumber(0, "=GCD(A7:A8)");
    }

    [Fact]
    public void Scalar_functions_apply_element_wise_to_arrays()
    {
        Assert.Equal(CellValue.Array(new CellValue[,] { { 1, 2 } }), Eval("=SQRT({1,4})"));
    }

    [Fact]
    public void RANDBETWEEN_is_volatile_and_uses_the_workbook_generator()
    {
        _wb.Random = new Random(7);
        var expected = Math.Floor(new Random(7).NextDouble() * 11) + 1;
        Assert.Equal(CellValue.Number(expected), Eval("=RANDBETWEEN(0.5,11.5)"));
        AssertError(ErrorKind.Num, "=RANDBETWEEN(3,2)");
        AssertNumber(5, "=RANDBETWEEN(5,5)");
    }
}
