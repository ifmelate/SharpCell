using SharpCell;
using static SharpCell.Tests.Functions.Financial;

namespace SharpCell.Tests.Functions;

public class FinancialSecuritiesTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public FinancialSecuritiesTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = true;
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    public static TheoryData<string, double, int> ExcelValues => new()
    {
        { "=ACCRINT(39508,39691,39569,0.1,1000,2,0)", 16.666666667, 9 },
        { $"=ACCRINT({D(2008, 3, 5)},39691,39569,0.1,1000,2,0,FALSE)", 15.555555556, 9 },
        { $"=ACCRINT({D(2008, 4, 5)},39691,39569,0.1,1000,2,0,TRUE)", 7.222222222, 9 },
        { $"=ACCRINT({D(2007, 3, 1)},{D(2008, 8, 31)},{D(2008, 5, 1)},0.1,1000,2,0,TRUE)", 116.944444444, 9 },
        { $"=ACCRINT({D(2007, 3, 1)},{D(2008, 8, 31)},{D(2008, 5, 1)},0.1,1000,2,0,FALSE)", 66.944444444, 9 },
        { "=ACCRINTM(39539,39614,0.1,1000,3)", 20.54794521, 8 },
        { $"=ACCRINTM({D(2008, 4, 1)},{D(2008, 6, 15)},0.1)", 20.555555556, 9 },
        { "=DISC(39472,39614,97.975,100,0)", 0.05207142857142884, 15 },
        { "=DISC(39472,39614,97.975,100,1)", 0.05219366197183125, 15 },
        { "=DISC(39472,39614,97.975,100,2.4)", 0.051338028169014345, 15 },
        { $"=INTRATE({D(2008, 2, 15)},{D(2008, 5, 15)},1000000,1014420,2)", 0.05768, 5 },
        { $"=RECEIVED({D(2008, 2, 15)},{D(2008, 5, 15)},1000000,0.0575,2)", 1014584.654, 3 },
        { $"=PRICEDISC({D(2008, 2, 16)},{D(2008, 3, 1)},0.0525,100,2)", 99.79583, 5 },
        { $"=YIELDDISC({D(2008, 2, 16)},{D(2008, 3, 1)},99.795,100,2)", 0.052823, 6 },
        { $"=PRICEMAT({D(2008, 2, 15)},{D(2008, 4, 13)},{D(2007, 11, 11)},0.061,0.061,0)", 99.98449888, 8 },
        { $"=YIELDMAT({D(2008, 3, 15)},{D(2008, 11, 3)},{D(2007, 11, 8)},0.0625,100.0123,0)", 0.060954, 6 },
        { $"=TBILLEQ({D(2008, 3, 31)},{D(2008, 6, 1)},0.0914)", 0.094151, 6 },
        { $"=TBILLPRICE({D(2008, 3, 31)},{D(2008, 6, 1)},0.09)", 98.45, 9 },
        { $"=TBILLYIELD({D(2008, 3, 31)},{D(2008, 6, 1)},98.45)", 0.091417, 6 },
    };

    [Theory]
    [MemberData(nameof(ExcelValues))]
    public void Matches_Excel(string formula, double expected, int decimals)
    {
        Rounds(expected, decimals, Eval(formula));
    }

    [Theory]
    // Basis outside 0 to 4, frequency other than 1, 2, 4, dates out of order.
    [InlineData("=DISC(39472,39614,97.975,100,5)", ErrorKind.Num)]
    [InlineData("=DISC(39472,39614,97.975,100,-1)", ErrorKind.Num)]
    [InlineData("=DISC(39614,39472,97.975,100)", ErrorKind.Num)]
    [InlineData("=DISC(39472,39472,97.975,100)", ErrorKind.Num)]
    [InlineData("=DISC(39472,39614,0,100)", ErrorKind.Num)]
    [InlineData("=DISC(-1,39614,97.975,100)", ErrorKind.Num)]
    [InlineData("=DISC(39472,3000000,97.975,100)", ErrorKind.Num)]
    [InlineData("=ACCRINT(39508,39691,39569,0.1,1000,3)", ErrorKind.Num)]
    [InlineData("=ACCRINT(39569,39691,39508,0.1,1000,2)", ErrorKind.Num)]
    [InlineData("=ACCRINT(39508,39691,39569,0,1000,2)", ErrorKind.Num)]
    [InlineData("=ACCRINT(39508,39691,39569,0.1,0,2)", ErrorKind.Num)]
    [InlineData("=ACCRINTM(39539,39614,0.1,1000,5)", ErrorKind.Num)]
    [InlineData("=RECEIVED(39472,39614,1000,3)", ErrorKind.Num)]
    [InlineData("=PRICEMAT(39500,39614,39600,0.06,0.06)", ErrorKind.Num)]
    [InlineData("=YIELDMAT(39500,39614,39400,0.06,0)", ErrorKind.Num)]
    [InlineData("=TBILLPRICE(39472,39472+367,0.09)", ErrorKind.Num)]
    [InlineData("=TBILLPRICE(39614,39472,0.09)", ErrorKind.Num)]
    [InlineData("=TBILLPRICE(39472,39614,0)", ErrorKind.Num)]
    [InlineData("=TBILLPRICE(39472,39472+360,3)", ErrorKind.Num)]
    [InlineData("=TBILLYIELD(39472,39614,0)", ErrorKind.Num)]
    [InlineData("=TBILLEQ(S!A1,39614,0.09)", ErrorKind.Value)]
    [InlineData("=DISC(39472,39614,97.975,TRUE)", ErrorKind.Value)]
    [InlineData("=ACCRINT(39508,39691,39569,0.1,1000,2,0,\"x\")", ErrorKind.Value)]
    public void Errors(string formula, ErrorKind expected)
    {
        Assert.Equal(Error(expected), Eval(formula));
    }

    [Fact]
    public void Dates_are_truncated_to_whole_days()
    {
        Assert.Equal(Eval("=DISC(39472,39614,97.975,100)"), Eval("=DISC(39472.9,39614.5,97.975,100)"));
    }

    [Fact]
    public void Dates_may_be_number_text()
    {
        Assert.Equal(Eval("=TBILLPRICE(39538,39600,0.09)"), Eval("=TBILLPRICE(\"39538\",\"39600\",0.09)"));
    }

    [Fact]
    public void Actual_actual_averages_the_years_over_a_long_span()
    {
        // 2008-01-01 to 2011-01-01: 1096 days over four calendar years of 1461 days.
        Near(1096 / (1461 / 4.0), Eval($"=ACCRINTM({D(2008, 1, 1)},{D(2011, 1, 1)},1,1,1)"));
    }
}
