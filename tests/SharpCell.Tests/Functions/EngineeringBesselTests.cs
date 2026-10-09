using System;
using SharpCell;

namespace SharpCell.Tests.Functions;

public class EngineeringBesselTests
{
    private readonly Workbook _wb = new();

    public EngineeringBesselTests()
    {
        var s = _wb.AddSheet("S");
        s["A1"].Value = true;
        s["A2"].Value = "1";
    }

    // Reference values to full precision (mpmath); Excel's own results agree
    // to about eight digits.
    [Theory]
    [InlineData("=BESSELJ(1,0)", 0.76519768655796655)]
    [InlineData("=BESSELJ(1,1)", 0.44005058574493352)]
    [InlineData("=BESSELJ(-1,3)", -0.019563353982668406)]
    [InlineData("=BESSELJ(10,0)", -0.24593576445134834)]
    [InlineData("=BESSELJ(10,10)", 0.20748610663335886)]
    [InlineData("=BESSELJ(32,7)", -0.076210348698464048)]
    [InlineData("=BESSELJ(0.2,10)", 2.7532277551302929e-17)]
    [InlineData("=BESSELJ(2,30)", 3.6502562664740971e-33)]
    [InlineData("=BESSELJ(50,100)", 1.1159273690838093e-21)]
    [InlineData("=BESSELJ(3000,2000)", 0.016448647918746624)]
    [InlineData("=BESSELJ(100000,5)", 0.001846551245452295)]
    [InlineData("=BESSELJ(10000000,3)", -0.00023689920557709944)]
    [InlineData("=BESSELY(1,0)", 0.088256964215676958)]
    [InlineData("=BESSELY(1,1)", -0.78121282130028872)]
    [InlineData("=BESSELY(10,0)", 0.055671167283599391)]
    [InlineData("=BESSELY(32,10)", -0.14177748068099433)]
    [InlineData("=BESSELY(1E-10,1)", -6366197723.6758132)]
    [InlineData("=BESSELY(2,30)", -2.9132238482189047e+30)]
    [InlineData("=BESSELY(50,100)", -3.2938001882026666e+18)]
    [InlineData("=BESSELY(3000,2000)", 0.0037612411453787259)]
    [InlineData("=BESSELY(100000,5)", 0.001719431949648651)]
    [InlineData("=BESSELY(10000000,3)", -8.6837455246570273e-5)]
    [InlineData("=BESSELI(1,0)", 1.2660658777520083)]
    [InlineData("=BESSELI(1,1)", 0.56515910399248503)]
    [InlineData("=BESSELI(-10,7)", -238.02558477578199)]
    [InlineData("=BESSELI(32,0)", 5590908381350.8731)]
    [InlineData("=BESSELI(0.2,10)", 2.7582381773426827e-17)]
    [InlineData("=BESSELI(1000,2000)", 1.2950953907756806e-285)]
    [InlineData("=BESSELK(1,0)", 0.42102443824070833)]
    [InlineData("=BESSELK(1,1)", 0.60190723019723457)]
    [InlineData("=BESSELK(0.2,10)", 1812385259400248.9)]
    [InlineData("=BESSELK(1E-10,0)", 23.141782445598869)]
    [InlineData("=BESSELK(1E-10,5)", 3.8399999999999993e+52)]
    [InlineData("=BESSELK(32,7)", 5.924792453809896e-15)]
    [InlineData("=BESSELK(2,30)", 4.2711257548876876e+30)]
    [InlineData("=BESSELK(700,1000)", 6.5156197914473582e-31)]
    [InlineData("=BESSELJ(A2,2)", 0.11490348493190048)]
    public void Matches_reference_values(string formula, double expected)
    {
        var actual = _wb.Evaluate(formula);
        Assert.Equal(CellValueKind.Number, actual.Kind);
        Assert.True(Math.Abs(actual.AsNumber() - expected) <= 1e-11 * Math.Abs(expected), $"{formula}: expected {expected}, got {actual}");
    }

    [Theory]
    [InlineData("=BESSELJ(0,0)", 1)]
    [InlineData("=BESSELJ(0,3)", 0)]
    [InlineData("=BESSELI(0,0)", 1)]
    [InlineData("=BESSELI(0,2)", 0)]
    [InlineData("=BESSELK(100000,5)", 0)]
    [InlineData("=BESSELJ(1,100000000)", 0)]
    [InlineData("=BESSELI(-3,100000000)", 0)]
    [InlineData("=BESSELJ(1E-300,2)", 0)]
    public void Exact_values(string formula, double expected) => Assert.Equal(CellValue.Number(expected), _wb.Evaluate(formula));

    [Theory]
    [InlineData("=BESSELJ(1,-1)", ErrorKind.Num)]
    [InlineData("=BESSELI(1,-1)", ErrorKind.Num)]
    [InlineData("=BESSELK(0,1)", ErrorKind.Num)]
    [InlineData("=BESSELK(-1,1)", ErrorKind.Num)]
    [InlineData("=BESSELY(0,0)", ErrorKind.Num)]
    [InlineData("=BESSELY(-2,1)", ErrorKind.Num)]
    [InlineData("=BESSELI(1000,1)", ErrorKind.Num)]
    [InlineData("=BESSELY(1,100000000)", ErrorKind.Num)]
    [InlineData("=BESSELY(1,1000)", ErrorKind.Num)]
    [InlineData("=BESSELK(1,1000)", ErrorKind.Num)]
    [InlineData("=BESSELJ(A1,1)", ErrorKind.Value)]
    [InlineData("=BESSELJ(1,A1)", ErrorKind.Value)]
    [InlineData("=BESSELY(\"x\",1)", ErrorKind.Value)]
    [InlineData("=BESSELK(1/0,1)", ErrorKind.Div0)]
    public void Errors(string formula, ErrorKind expected) => Assert.Equal(CellValue.Error(expected), _wb.Evaluate(formula));

    [Fact]
    public void The_order_is_truncated()
    {
        Assert.Equal(_wb.Evaluate("=BESSELY(1,1)"), _wb.Evaluate("=BESSELY(1,1.999)"));
    }
}
