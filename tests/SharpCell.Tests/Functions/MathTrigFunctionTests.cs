using System;
using SharpCell;

namespace SharpCell.Tests.Functions;

public class MathTrigFunctionTests
{
    private readonly Workbook _wb = new();

    public MathTrigFunctionTests() => _wb.AddSheet("S");

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    [Theory]
    [InlineData("=SIN(1)", 0.8414709848078965)]
    [InlineData("=COS(5)", 0.28366218546322625)]
    [InlineData("=TAN(-10)", -0.6483608274590866)]
    [InlineData("=SIN(2^20)", 0.33049314002171887)]
    [InlineData("=COT(1)", 0.6420926159343306)]
    [InlineData("=CSC(1)", 1.1883951057781212)]
    [InlineData("=SEC(1)", 1.8508157176809255)]
    [InlineData("=ASIN(1)", Math.PI / 2)]
    [InlineData("=ACOS(-1)", Math.PI)]
    [InlineData("=ATAN(1)", Math.PI / 4)]
    [InlineData("=ATAN2(-1,-1)", -2.356194490192345)]
    [InlineData("=ACOT(-1)", 2.356194490192345)]
    [InlineData("=ACOT(0)", Math.PI / 2)]
    [InlineData("=ACOT(9999999999999)", 1.0000000000001E-13)]
    [InlineData("=SINH(-10)", -11013.232874703393)]
    [InlineData("=COSH(5)", 74.20994852478785)]
    [InlineData("=TANH(2.8)", 0.9926315202011279)]
    [InlineData("=COTH(1)", 1.3130352854993315)]
    [InlineData("=CSCH(-10)", -9.079985971212217E-05)]
    [InlineData("=SECH(1E+300)", 0)]
    [InlineData("=ASINH(-10)", -2.99822295029797)]
    [InlineData("=ACOSH(5)", 2.2924316695611777)]
    [InlineData("=ATANH(0.5)", 0.5493061443340549)]
    [InlineData("=ACOTH(-10)", -0.1003353477310756)]
    [InlineData("=ACOTH(9999999999999)", 1.0000000000001E-13)]
    [InlineData("=DEGREES(PI())", 180)]
    [InlineData("=RADIANS(180)", Math.PI)]
    public void Computes_like_Excel(string formula, double expected)
    {
        var actual = Eval(formula);
        Assert.Equal(CellValueKind.Number, actual.Kind);
        Assert.Equal(expected, actual.AsNumber(), 1e-12 * Math.Max(1, Math.Abs(expected)));
    }

    // Excel reduces angles by the x87 pi, so these differ from the exact values in the third digit.
    [Theory]
    [InlineData("=SIN(PI())", 1.22514845490862E-16)]
    [InlineData("=COS(PI()/2)", 6.1257422745431E-17)]
    [InlineData("=SIN(2*PI())", -2.45029690981724E-16)]
    [InlineData("=SIN(PI()-0.000000000001)", 1.0002114154278319E-12)]
    [InlineData("=TAN(PI()/2-0.000000000001)", 999849864538.9548)]
    [InlineData("=ASINH(1E-12)", 1.000088900581841E-12)]
    public void Matches_Excel_near_zeros_and_poles(string formula, double expected)
    {
        var actual = Eval(formula).AsNumber();
        Assert.True(Math.Abs(actual - expected) <= 1e-12 * Math.Abs(expected), $"{formula}: {actual:R}");
    }

    [Theory]
    [InlineData("=SIN(2^27)", ErrorKind.Num)]
    [InlineData("=COT(0)", ErrorKind.Div0)]
    [InlineData("=CSC(0)", ErrorKind.Div0)]
    [InlineData("=ASIN(1.5)", ErrorKind.Num)]
    [InlineData("=ACOS(-2)", ErrorKind.Num)]
    [InlineData("=ATAN2(0,0)", ErrorKind.Div0)]
    [InlineData("=SINH(1000)", ErrorKind.Num)]
    [InlineData("=COTH(0)", ErrorKind.Div0)]
    [InlineData("=CSCH(0)", ErrorKind.Div0)]
    [InlineData("=ACOSH(0.5)", ErrorKind.Num)]
    [InlineData("=ATANH(1)", ErrorKind.Num)]
    [InlineData("=ACOTH(1)", ErrorKind.Num)]
    [InlineData("=SIN(\"x\")", ErrorKind.Value)]
    public void Fails_like_Excel(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }
}
