using System;
using SharpCell;

namespace SharpCell.Tests.Functions;

public class MathRoundingFunctionTests
{
    private readonly Workbook _wb = new();

    public MathRoundingFunctionTests() => _wb.AddSheet("S");

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    [Theory]
    [InlineData("=ROUND(2.675,2)", 2.68)]
    [InlineData("=ROUND(1.005,2)", 1.01)]
    [InlineData("=ROUND(-2.5,0)", -3)]
    [InlineData("=ROUND(1234.5678,-2)", 1200)]
    [InlineData("=ROUND(1234.5678,1.9)", 1234.6)]
    [InlineData("=ROUND(5,-1)", 10)]
    [InlineData("=ROUND(4,-1)", 0)]
    [InlineData("=ROUND(1E+20,2)", 1E+20)]
    [InlineData("=ROUND(0.1+0.2,15)", 0.3)]
    [InlineData("=ROUNDUP(0.1*3,1)", 0.3)]
    [InlineData("=ROUNDUP(-3.14159,2)", -3.15)]
    [InlineData("=ROUNDUP(31415.92654,-2)", 31500)]
    [InlineData("=ROUNDDOWN(-3.14159,1)", -3.1)]
    [InlineData("=ROUNDDOWN(0.3/0.1,0)", 3)]
    [InlineData("=TRUNC(-8.9)", -8)]
    [InlineData("=TRUNC(PI(),1)", 3.1)]
    [InlineData("=INT(-7.5)", -8)]
    [InlineData("=EVEN(1.5)", 2)]
    [InlineData("=EVEN(-1)", -2)]
    [InlineData("=EVEN(0)", 0)]
    [InlineData("=ODD(0)", 1)]
    [InlineData("=ODD(2)", 3)]
    [InlineData("=ODD(-1.5)", -3)]
    [InlineData("=MROUND(10,3)", 9)]
    [InlineData("=MROUND(-7.5,-5)", -10)]
    [InlineData("=MROUND(2.675,0.01)", 2.68)]
    [InlineData("=MROUND(5,0)", 0)]
    [InlineData("=CEILING(2.5,1)", 3)]
    [InlineData("=CEILING(-2.5,1)", -2)]
    [InlineData("=CEILING(-2.5,-1)", -3)]
    [InlineData("=CEILING(0.3,0.1)", 0.3)]
    [InlineData("=CEILING(5,0)", 0)]
    [InlineData("=FLOOR(10,0.21)", 9.87)]
    [InlineData("=FLOOR(-10,7)", -14)]
    [InlineData("=FLOOR(-10,-7)", -7)]
    [InlineData("=FLOOR(0.3,0.1)", 0.3)]
    [InlineData("=FLOOR(0,0)", 0)]
    [InlineData("=CEILING.MATH(-5.5,2)", -4)]
    [InlineData("=CEILING.MATH(-5.5,2,-1)", -6)]
    [InlineData("=CEILING.MATH(6.7)", 7)]
    [InlineData("=FLOOR.MATH(-5.5,2)", -6)]
    [InlineData("=FLOOR.MATH(-5.5,2,1)", -4)]
    [InlineData("=FLOOR.MATH(5.5,-2)", 4)]
    [InlineData("=CEILING.PRECISE(-4.3,-2)", -4)]
    [InlineData("=ISO.CEILING(4.3)", 5)]
    [InlineData("=FLOOR.PRECISE(-3.2,-1)", -4)]
    [InlineData("=FLOOR.PRECISE(3.2,0)", 0)]
    public void Rounds_like_Excel(string formula, double expected)
    {
        var actual = Eval(formula);
        Assert.Equal(CellValueKind.Number, actual.Kind);
        Assert.Equal(expected, actual.AsNumber(), 1e-12 * Math.Max(1, Math.Abs(expected)));
    }

    [Theory]
    [InlineData("=MROUND(-7.5,5)", ErrorKind.Num)]
    [InlineData("=MROUND(TRUE,2)", ErrorKind.Value)]
    [InlineData("=MROUND(,2)", ErrorKind.NA)]
    [InlineData("=CEILING(2.5,-1)", ErrorKind.Num)]
    [InlineData("=FLOOR(10,0)", ErrorKind.Div0)]
    [InlineData("=FLOOR(10,-1)", ErrorKind.Num)]
    [InlineData("=ROUNDUP(5,-400)", ErrorKind.Num)]
    [InlineData("=ROUND(\"x\",1)", ErrorKind.Value)]
    public void Fails_like_Excel(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }
}
