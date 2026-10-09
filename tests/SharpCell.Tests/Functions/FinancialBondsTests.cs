using System.Globalization;
using SharpCell;
using static SharpCell.Tests.Functions.Financial;

namespace SharpCell.Tests.Functions;

public class FinancialBondsTests
{
    private readonly Workbook _wb = new();

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    private static readonly string Settlement = D(2011, 1, 25);
    private static readonly string Maturity = D(2011, 11, 15);

    public static TheoryData<string, double, int> ExcelValues => new()
    {
        { $"=COUPDAYBS({Settlement},{Maturity},2,1)", 71, 9 },
        { $"=COUPDAYS({Settlement},{Maturity},2,1)", 181, 9 },
        { $"=COUPDAYSNC({Settlement},{Maturity},2,1)", 110, 9 },
        { $"=COUPNCD({Settlement},{Maturity},2,1)", 40678, 9 },
        { $"=COUPNUM({Settlement},{Maturity},2,1)", 2, 9 },
        { $"=COUPPCD({Settlement},{Maturity},2,1)", 40497, 9 },
        { $"=COUPDAYS({Settlement},{Maturity},2,0)", 180, 9 },
        { $"=COUPDAYS({Settlement},{Maturity},4,3)", 91.25, 9 },
        { $"=COUPDAYBS({Settlement},{Maturity},2,0)", 70, 9 },
        { $"=COUPDAYSNC({Settlement},{Maturity},2,0)", 110, 9 },
        // A maturity on a month end keeps every coupon date on one.
        { $"=COUPPCD({D(2011, 1, 25)},{D(2011, 11, 30)},2)", 40512, 9 },
        { $"=COUPNCD({D(2011, 1, 25)},{D(2011, 11, 30)},2)", 40694, 9 },
        { $"=COUPNCD({D(2011, 1, 25)},{D(2011, 11, 30)},4)", 40602, 9 },
        { $"=PRICE({D(2008, 2, 15)},{D(2017, 11, 15)},0.0575,0.065,100,2,0)", 94.63436, 5 },
        { $"=YIELD({D(2008, 2, 15)},{D(2016, 11, 15)},0.0575,95.04287,100,2,0)", 0.065, 6 },
        { $"=DURATION({D(2018, 7, 1)},{D(2048, 1, 1)},0.08,0.09,2,1)", 10.9191453, 7 },
        { $"=MDURATION({D(2008, 1, 1)},{D(2016, 1, 1)},0.08,0.09,2,1)", 5.73567, 5 },
        { $"=ODDFPRICE({D(2008, 11, 11)},{D(2021, 3, 1)},{D(2008, 10, 15)},{D(2009, 3, 1)},0.0785,0.0625,100,2,1)", 113.597717, 6 },
        { $"=ODDFYIELD({D(2008, 11, 11)},{D(2021, 3, 1)},{D(2008, 10, 15)},{D(2009, 3, 1)},0.0575,84.5,100,2,0)", 0.0772, 4 },
        { $"=ODDLPRICE({D(2008, 2, 7)},{D(2008, 6, 15)},{D(2007, 10, 15)},0.0375,0.0405,100,2,0)", 99.87829, 5 },
        { $"=ODDLYIELD({D(2008, 4, 20)},{D(2008, 6, 15)},{D(2007, 12, 24)},0.0375,99.875,100,2,0)", 0.045192, 6 },
    };

    [Theory]
    [MemberData(nameof(ExcelValues))]
    public void Matches_Excel(string formula, double expected, int decimals)
    {
        Rounds(expected, decimals, Eval(formula));
    }

    [Theory]
    [InlineData("=COUPDAYS(40568,40862,3)", ErrorKind.Num)]
    [InlineData("=COUPDAYS(40568,40862,2,5)", ErrorKind.Num)]
    [InlineData("=COUPDAYS(40862,40568,2)", ErrorKind.Num)]
    [InlineData("=COUPDAYS(40568,40568,2)", ErrorKind.Num)]
    [InlineData("=COUPNUM(TRUE,40862,2)", ErrorKind.Value)]
    [InlineData("=COUPNUM(\"x\",40862,2)", ErrorKind.Value)]
    [InlineData("=PRICE(39493,43054,-0.01,0.065,100,2)", ErrorKind.Num)]
    [InlineData("=PRICE(39493,43054,0.0575,0.065,0,2)", ErrorKind.Num)]
    [InlineData("=YIELD(39493,43054,0.0575,0,100,2)", ErrorKind.Num)]
    [InlineData("=DURATION(39493,43054,0.08,-0.09,2)", ErrorKind.Num)]
    [InlineData("=ODDFPRICE(39763,44256,39740,39873,0.0785,0.0625,100,3)", ErrorKind.Num)]
    // Issue after the first coupon; first coupon after maturity.
    [InlineData("=ODDFPRICE(39763,44256,39880,39873,0.0785,0.0625,100,2)", ErrorKind.Num)]
    [InlineData("=ODDFPRICE(39763,39800,39740,39873,0.0785,0.0625,100,2)", ErrorKind.Num)]
    [InlineData("=ODDLPRICE(39485,39614,39614,0.0375,0.0405,100,2)", ErrorKind.Num)]
    [InlineData("=ODDLYIELD(39558,39614,39440,0.0375,0,100,2)", ErrorKind.Num)]
    public void Errors(string formula, ErrorKind expected)
    {
        Assert.Equal(Error(expected), Eval(formula));
    }

    [Fact]
    public void Coupon_dates_follow_the_date_system()
    {
        var wb = new Workbook { DateSystem = DateSystem.Date1904 };
        // 40678 in the 1900 system is 39216 in the 1904 system (1462 days apart).
        Assert.Equal(CellValue.Number(40678 - 1462), wb.Evaluate($"=COUPNCD({40568 - 1462},{40862 - 1462},2,1)"));
    }

    [Fact]
    public void Yield_inverts_price()
    {
        var price = Eval($"=PRICE({D(2010, 3, 1)},{D(2030, 9, 30)},0.04,0.11,105,4,4)").AsNumber().ToString("R", CultureInfo.InvariantCulture);
        Near(0.11, Eval($"=YIELD({D(2010, 3, 1)},{D(2030, 9, 30)},0.04,{price},105,4,4)"));
    }

    [Fact]
    public void One_coupon_left_is_priced_with_simple_interest()
    {
        // 90 of 180 days to the coupon of 2.875 paid with the redemption.
        Near(102.875 / (1 + 0.5 * 0.065 / 2) - 0.5 * 2.875, Eval($"=PRICE({D(2008, 2, 15)},{D(2008, 5, 15)},0.0575,0.065,100,2,0)"));
        var price = (102.875 / (1 + 0.5 * 0.065 / 2) - 0.5 * 2.875).ToString("R", CultureInfo.InvariantCulture);
        Near(0.065, Eval($"=YIELD({D(2008, 2, 15)},{D(2008, 5, 15)},0.0575,{price},100,2,0)"));
    }
}
