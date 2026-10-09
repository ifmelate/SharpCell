using SharpCell;
using static SharpCell.Tests.Functions.Financial;

namespace SharpCell.Tests.Functions;

public class FinancialCashFlowsTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public FinancialCashFlowsTests()
    {
        _s = _wb.AddSheet("S");
        // A: a project; B: the same with text, a logical value and an empty cell in it.
        double[] flows = [-70000, 12000, 15000, 18000, 21000, 26000];
        for (var i = 0; i < flows.Length; i++)
            _s[i + 1, 1].Value = flows[i];
        _s["B1"].Value = -70000;
        _s["B2"].Value = "x";
        _s["B3"].Value = 12000;
        _s["B4"].Value = true;
        _s["B6"].Value = 15000;
        // XNPV and XIRR: values in C, dates in D.
        double[] values = [-10000, 2750, 4250, 3250, 2750];
        string[] dates = [D(2008, 1, 1), D(2008, 3, 1), D(2008, 10, 30), D(2009, 2, 15), D(2009, 4, 1)];
        for (var i = 0; i < values.Length; i++)
        {
            _s[i + 1, 3].Value = values[i];
            _s[i + 1, 4].Formula = "=" + dates[i];
        }

        _wb.Recalculate();
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    [Theory]
    [InlineData("=NPV(0.1,-10000,3000,4200,6800)", 1188.44, 2)]
    [InlineData("=NPV(0.1,\"10\",TRUE)", 9.917355372, 9)]
    [InlineData("=IRR(S!A1:A5)", -0.021244848, 9)]
    [InlineData("=IRR(S!A1:A6)", 0.086630948, 9)]
    [InlineData("=IRR(S!A1:A3,-0.1)", -0.443506941, 9)]
    [InlineData("=IRR({-20,5,5,-10,5})", -0.448928350, 9)]
    [InlineData("=IRR(S!A:A)", 0.086630948, 9)]
    [InlineData("=MIRR({-120000,39000,30000,21000,37000,46000},0.1,0.12)", 0.126094, 6)]
    [InlineData("=MIRR({-120000,39000,30000,21000},0.1,0.12)", -0.048044655, 9)]
    [InlineData("=MIRR({-120000,39000,30000,21000,37000,46000},0.1,0.14)", 0.134759111, 9)]
    [InlineData("=XNPV(0.09,S!C1:C5,S!D1:D5)", 2086.647602, 6)]
    [InlineData("=XIRR(S!C1:C5,S!D1:D5)", 0.37336253, 8)]
    [InlineData("=XIRR(S!C1:C5,S!D1:D5,0.5)", 0.37336253, 8)]
    [InlineData("=FVSCHEDULE(1,{0.09,0.11,0.1})", 1.33089, 9)]
    [InlineData("=FVSCHEDULE(1000,0.05)", 1050, 9)]
    public void Matches_Excel(string formula, double expected, int decimals)
    {
        Rounds(expected, decimals, Eval(formula));
    }

    // -70000, 12000 and 15000 in periods 1 to 3: text and logical values in a range are skipped,
    // and so are empty cells.
    [Fact]
    public void Npv_takes_only_the_numbers_of_a_range()
    {
        Near(-70000 / 1.1 + 12000 / (1.1 * 1.1) + 15000 / (1.1 * 1.1 * 1.1), Eval("=NPV(0.1,S!B1:B6)"));
    }

    [Theory]
    [InlineData("=NPV(0.1,1,#DIV/0!)", ErrorKind.Div0)]
    [InlineData("=NPV(0.1,\"x\")", ErrorKind.Value)]
    [InlineData("=IRR({1,2,3})", ErrorKind.Num)]
    [InlineData("=IRR({-1,-2})", ErrorKind.Num)]
    [InlineData("=IRR(S!A1)", ErrorKind.Num)]
    [InlineData("=IRR(S!A1:A6,-1)", ErrorKind.Value)]
    [InlineData("=IRR(S!A1:A6,\"x\")", ErrorKind.Value)]
    [InlineData("=MIRR({5,6},0.1,0.2)", ErrorKind.Div0)]
    [InlineData("=XNPV(0.09,S!C1:C4,S!D1:D5)", ErrorKind.Num)]
    [InlineData("=XNPV(0,S!C1:C5,S!D1:D5)", ErrorKind.Num)]
    [InlineData("=XNPV(0.09,{1,\"x\"},{1,2})", ErrorKind.Num)]
    [InlineData("=XNPV(0.09,{1,2},{5,4})", ErrorKind.Num)]
    [InlineData("=XNPV(0.09,{1,2},{-1,4})", ErrorKind.Num)]
    [InlineData("=XIRR({-1,\"x\"},{1,2})", ErrorKind.Value)]
    [InlineData("=XIRR({1,2},{1,2})", ErrorKind.Num)]
    [InlineData("=XIRR(S!C1:C5,S!D1:D5,-1)", ErrorKind.Value)]
    [InlineData("=FVSCHEDULE(1,{0.1,\"x\"})", ErrorKind.Value)]
    [InlineData("=FVSCHEDULE(TRUE,0.1)", ErrorKind.Value)]
    public void Errors(string formula, ErrorKind expected)
    {
        Assert.Equal(Error(expected), Eval(formula));
    }

    [Fact]
    public void Npv_skips_logical_values_in_single_cell_references()
    {
        _s["E1"].Value = true;
        _s["E2"].Value = 7;
        Near(7 / 1.07, Eval("=NPV(0.07,S!E1,S!E2)"));
    }

    [Fact]
    public void Xirr_counts_empty_values_as_zero_but_xnpv_refuses_them()
    {
        // F has an empty cell where I has 0; G holds the dates.
        double[] values = [-20, 5, 0, 15];
        for (var row = 1; row <= 4; row++)
        {
            if (values[row - 1] != 0)
                _s[row, 6].Value = values[row - 1];
            _s[row, 9].Value = values[row - 1];
            _s[row, 7].Value = 40000 + row * 100;
        }

        _s["H1"].Formula = "=XIRR(F1:F4,G1:G4)";
        _s["H2"].Formula = "=XIRR(I1:I4,G1:G4)";
        _s["H3"].Formula = "=XNPV(0.1,F1:F4,G1:G4)";
        _wb.Recalculate();
        Assert.Equal(CellValueKind.Number, _s["H1"].Value.Kind);
        Assert.Equal(_s["H2"].Value, _s["H1"].Value);
        Assert.Equal(Error(ErrorKind.Num), _s["H3"].Value);
    }

    [Fact]
    public void Xirr_finds_a_rate_below_minus_ninety_percent()
    {
        // Excel's own result is accurate to about 1e-8, the tolerance its corpus file carries.
        Near(5e-8,-0.9977533546974883, Eval($"=XIRR({{-400,100,10,5,6,5}},{{{D(2023, 2, 1)},{D(2023, 3, 1)},{D(2023, 4, 1)},{D(2023, 6, 1)},{D(2023, 7, 1)},{D(2023, 8, 1)}}})"));
    }
}
