using SharpCell;
using static SharpCell.Tests.Functions.Financial;

namespace SharpCell.Tests.Functions;

public class FinancialDepreciationTests
{
    private readonly Workbook _wb = new();

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    public static TheoryData<string, double, int> ExcelValues => new()
    {
        { "=SLN(30000,7500,10)", 2250, 9 },
        { "=SYD(30000,7500,10,1)", 4090.91, 2 },
        { "=SYD(30000,7500,10,10)", 409.09, 2 },
        { "=SYD(100,5,5.5,5.4)", 5.846153846, 9 },
        { "=DB(1000000,100000,6,1,7)", 186083.33, 2 },
        { "=DB(1000000,100000,6,2,7)", 259639.42, 2 },
        { "=DB(1000000,100000,6,3,7)", 176814.44, 2 },
        { "=DB(1000000,100000,6,4,7)", 120410.64, 2 },
        { "=DB(1000000,100000,6,5,7)", 81999.64, 2 },
        { "=DB(1000000,100000,6,6,7)", 55841.76, 2 },
        { "=DB(1000000,100000,6,7,7)", 15845.10, 2 },
        { "=DB(100,10,4,3,11)", 14.7324366, 7 },
        { "=DB(100,10,4,3)", 13.8339672, 7 },
        { "=DDB(2400,300,10*365,1)", 1.32, 2 },
        { "=DDB(2400,300,10*12,1,2)", 40.00, 2 },
        { "=DDB(2400,300,10,1,2)", 480, 9 },
        { "=DDB(2400,300,10,2,1.5)", 306, 9 },
        { "=DDB(2400,300,10,10)", 22.12, 2 },
        { "=DDB(100,10,4,3,40)", 0, 9 },
        { "=VDB(2400,300,10*365,0,1)", 1.32, 2 },
        { "=VDB(2400,300,10*12,0,1)", 40.00, 2 },
        { "=VDB(2400,300,10,0,1)", 480, 9 },
        { "=VDB(2400,300,10*12,6,18)", 396.31, 2 },
        { "=VDB(2400,300,10*12,6,18,1.5)", 311.81, 2 },
        { "=VDB(2400,300,10,0,0.875,1.5)", 315, 9 },
        { "=VDB(2400,300,10,0,10)", 2100, 9 },
        // From the fourth year straight line (1080 a year) beats the declining balance (864),
        // unless switching is off.
        { "=VDB(10000,0,5,3,4)", 1080, 9 },
        { "=VDB(10000,0,5,4,5)", 1080, 9 },
        { "=VDB(10000,0,5,3,4,2,TRUE)", 864, 9 },
        { "=VDB(10000,0,5,4,5,2,TRUE)", 518.4, 9 },
        { $"=AMORDEGRC(2400,{D(2008, 8, 19)},{D(2008, 12, 31)},300,1,0.15,1)", 776, 9 },
        // Period 0 is the prorated first period: 134 of 366 days at 0.15, or 0.15·2.5 rounded.
        { $"=AMORDEGRC(2400,{D(2008, 8, 19)},{D(2008, 12, 31)},300,0,0.15,1)", 330, 9 },
        { $"=AMORLINC(2400,{D(2008, 8, 19)},{D(2008, 12, 31)},300,1,0.15,1)", 360, 9 },
        { $"=AMORLINC(2400,{D(2008, 8, 19)},{D(2008, 12, 31)},300,0,0.15,1)", 131.8032787, 7 },
        { $"=AMORLINC(2400,{D(2008, 8, 19)},{D(2008, 12, 31)},300,9,0.15,1)", 0, 9 },
    };

    [Theory]
    [MemberData(nameof(ExcelValues))]
    public void Matches_Excel(string formula, double expected, int decimals)
    {
        Rounds(expected, decimals, Eval(formula));
    }

    [Theory]
    [InlineData("=SLN(100,10,0)", ErrorKind.Div0)]
    [InlineData("=SYD(100,10,5,6)", ErrorKind.Num)]
    [InlineData("=SYD(100,10,5,0)", ErrorKind.Num)]
    [InlineData("=DB(100,10,4,5)", ErrorKind.Num)]
    [InlineData("=DB(100,10,4,6,2)", ErrorKind.Num)]
    [InlineData("=DB(100,10,4,3,13)", ErrorKind.Num)]
    [InlineData("=DB(100,10,4,3,0)", ErrorKind.Num)]
    [InlineData("=DB(-100,10,4,3)", ErrorKind.Num)]
    [InlineData("=DDB(100,10,4,5)", ErrorKind.Num)]
    [InlineData("=DDB(100,-10,4,3)", ErrorKind.Num)]
    [InlineData("=DDB(100,10,4,3,0)", ErrorKind.Num)]
    [InlineData("=VDB(2400,300,10,5,4)", ErrorKind.Num)]
    [InlineData("=VDB(2400,300,10,0,11)", ErrorKind.Num)]
    [InlineData("=VDB(2400,300,0,0,1)", ErrorKind.Num)]
    [InlineData("=VDB(2400,300,10,10,10)", ErrorKind.Num)]
    [InlineData("=AMORLINC(2400,39679,39813,300,1,0.15,2)", ErrorKind.Num)]
    [InlineData("=AMORLINC(2400,39813,39679,300,1,0.15,1)", ErrorKind.Num)]
    [InlineData("=AMORLINC(2400,39679,39813,2400,1,0.15,1)", ErrorKind.Num)]
    [InlineData("=AMORLINC(2400,39679,39813,300,1,0,1)", ErrorKind.Num)]
    // A life of 4 to 5 years (rate 0.2 to 0.25) has no degressive coefficient.
    [InlineData("=AMORDEGRC(2400,39679,39813,300,1,0.22,1)", ErrorKind.Num)]
    [InlineData("=AMORDEGRC(2400,39679,39813,300,1,0.4,1)", ErrorKind.Num)]
    [InlineData("=AMORDEGRC(TRUE,39679,39813,300,1,0.15,1)", ErrorKind.Value)]
    public void Errors(string formula, ErrorKind expected)
    {
        Assert.Equal(Error(expected), Eval(formula));
    }

    [Fact]
    public void Db_rounds_the_rate_to_three_decimals()
    {
        // 1 − (10/100)^(1/4) = 0.43766, used as 0.438.
        Near(100 * 0.438, Eval("=DB(100,10,4,1)"));
    }
}
