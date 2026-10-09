using System;
using System.Globalization;
using SharpCell;

namespace SharpCell.Tests.Functions;

/// <summary>Helpers for the financial tests: date serials (DATE belongs to another category) and rounded comparisons.</summary>
internal static class Financial
{
    /// <summary>The 1900-system serial number of a date, as text for a formula.</summary>
    public static string D(int year, int month, int day)
    {
        Assert.True(DateSerial.TryFromDate(year, month, day, DateSystem.Date1900, out var serial));
        return serial.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>The value equals <paramref name="expected"/> as shown with <paramref name="decimals"/> decimals.</summary>
    public static void Rounds(double expected, int decimals, CellValue actual)
    {
        Assert.Equal(CellValueKind.Number, actual.Kind);
        Assert.Equal(expected, Math.Round(actual.AsNumber(), decimals, MidpointRounding.AwayFromZero), 12);
    }

    /// <summary>The value is within a relative 1e-9 of <paramref name="expected"/>, the corpus tolerance.</summary>
    public static void Near(double expected, CellValue actual) => Near(1e-9, expected, actual);

    public static void Near(double tolerance, double expected, CellValue actual)
    {
        Assert.Equal(CellValueKind.Number, actual.Kind);
        var difference = Math.Abs(actual.AsNumber() - expected);
        Assert.True(difference <= tolerance * Math.Max(Math.Abs(expected), Math.Abs(actual.AsNumber())),
            $"expected {expected:R}, got {actual.AsNumber():R}");
    }

    public static CellValue Error(ErrorKind kind) => CellValue.Error(kind);
}

public class FinancialFunctionsTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public FinancialFunctionsTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = true;
        _s["A2"].Value = "text";
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    [Theory]
    [InlineData("=PV(0.08/12,12*20,500,,0)", -59777.15, 2)]
    [InlineData("=PV(0,10,-100,-50)", 1050, 9)]
    [InlineData("=FV(0.06/12,10,-200,-500,1)", 2581.40, 2)]
    [InlineData("=FV(0.12/12,12,-1000)", 12682.50, 2)]
    [InlineData("=FV(0,3,100,300,1)", -600, 9)]
    [InlineData("=PMT(0.08/12,10,10000)", -1037.03, 2)]
    [InlineData("=PMT(0.06/12,18*12,0,50000)", -129.08, 2)]
    [InlineData("=PMT(0,3,100,300,1)", -133.333333333, 9)]
    [InlineData("=NPER(0.12/12,-100,-1000,10000,1)", 59.6738657, 7)]
    [InlineData("=NPER(0.12/12,-100,-1000)", -9.57859404, 8)]
    [InlineData("=NPER(0,-100,1000)", 10, 9)]
    [InlineData("=RATE(4*12,-200,8000)", 0.007701472, 9)]
    [InlineData("=RATE(48,-200,8000,0,1)", 0.008052982, 9)]
    [InlineData("=IPMT(0.1/12,1,3*12,8000)", -66.67, 2)]
    [InlineData("=IPMT(0.1,3,3,8000)", -292.45, 2)]
    [InlineData("=IPMT(0.1,1,3,8000,0,1)", 0, 9)]
    [InlineData("=PPMT(0.1/12,1,2*12,2000)", -75.62, 2)]
    [InlineData("=PPMT(0.08,10,10,200000)", -27598.05, 2)]
    [InlineData("=CUMIPMT(0.09/12,30*12,125000,13,24,0)", -11135.23213, 5)]
    [InlineData("=CUMIPMT(0.09/12,30*12,125000,1,1,0)", -937.5, 9)]
    [InlineData("=CUMPRINC(0.09/12,30*12,125000,13,24,0)", -934.1071234, 7)]
    [InlineData("=CUMPRINC(0.09/12,30*12,125000,1,1,0)", -68.27827118, 8)]
    [InlineData("=ISPMT(0.1/12,1,3*12,8000000)", -64814.81, 2)]
    [InlineData("=ISPMT(0.1,1,3,8000000)", -533333.33, 2)]
    [InlineData("=PDURATION(0.025,2000,2200)", 3.86, 2)]
    [InlineData("=PDURATION(0.025/12,1000,1200)", 87.6, 1)]
    [InlineData("=RRI(96,10000,11000)", 0.0009933, 7)]
    [InlineData("=EFFECT(0.0525,4)", 0.05354267, 8)]
    [InlineData("=EFFECT(0.0525,4.9)", 0.05354267, 8)]
    [InlineData("=NOMINAL(0.053543,4)", 0.05250032, 8)]
    [InlineData("=DOLLARDE(1.02,16)", 1.125, 9)]
    [InlineData("=DOLLARDE(1.1,32)", 1.3125, 9)]
    [InlineData("=DOLLARDE(-1.02,16)", -1.125, 9)]
    [InlineData("=DOLLARDE(1.1,10)", 1.1, 9)]
    [InlineData("=DOLLARFR(1.125,16)", 1.02, 9)]
    [InlineData("=DOLLARFR(1.125,32)", 1.04, 9)]
    public void Matches_Excel(string formula, double expected, int decimals)
    {
        Financial.Rounds(expected, decimals, Eval(formula));
    }

    [Theory]
    // Rate −100% or below: PMT cannot be found; PV divides by zero.
    [InlineData("=PMT(-1,3,100)", ErrorKind.Num)]
    [InlineData("=PMT(0,0,100)", ErrorKind.Num)]
    [InlineData("=PV(-1,3,100)", ErrorKind.Div0)]
    [InlineData("=FV(-1,-2,1)", ErrorKind.Div0)]
    [InlineData("=FV(-3,-3.5,1)", ErrorKind.Num)]
    [InlineData("=NPER(0,0,100)", ErrorKind.Div0)]
    [InlineData("=NPER(0.02,0,100,5,1)", ErrorKind.Num)]
    [InlineData("=NPER(-2,-100,1000)", ErrorKind.Num)]
    [InlineData("=PV(\"x\",3,100)", ErrorKind.Value)]
    [InlineData("=PMT(0.1,3,#N/A)", ErrorKind.NA)]
    // Newton's method from 0.1 runs away for these.
    [InlineData("=RATE(10,-1200,2000,0,1)", ErrorKind.Num)]
    [InlineData("=RATE(3300,-200,8000)", ErrorKind.Num)]
    [InlineData("=RATE(48,-200,8000,0,0,-1)", ErrorKind.Value)]
    [InlineData("=IPMT(0.1,0,3,8000)", ErrorKind.Num)]
    [InlineData("=IPMT(0.1,4,3,8000)", ErrorKind.Num)]
    [InlineData("=CUMIPMT(0.1,36,8000,0,10,0)", ErrorKind.Num)]
    [InlineData("=CUMIPMT(0.1,36,8000,5,4,0)", ErrorKind.Num)]
    [InlineData("=CUMIPMT(0.1,36,8000,1,10,2)", ErrorKind.Num)]
    [InlineData("=CUMIPMT(0,36,8000,1,10,0)", ErrorKind.Num)]
    [InlineData("=CUMPRINC(0.1,36,-8000,1,10,0)", ErrorKind.Num)]
    [InlineData("=ISPMT(0.1,1,0,100)", ErrorKind.Div0)]
    [InlineData("=PDURATION(0,100,200)", ErrorKind.Num)]
    [InlineData("=PDURATION(0.1,-100,200)", ErrorKind.Num)]
    [InlineData("=RRI(0,100,200)", ErrorKind.Num)]
    [InlineData("=RRI(2,0,200)", ErrorKind.Num)]
    [InlineData("=EFFECT(0,4)", ErrorKind.Num)]
    [InlineData("=EFFECT(0.1,0.9)", ErrorKind.Num)]
    [InlineData("=NOMINAL(-0.1,4)", ErrorKind.Num)]
    [InlineData("=DOLLARDE(1.02,0.5)", ErrorKind.Div0)]
    [InlineData("=DOLLARDE(1.02,-1)", ErrorKind.Num)]
    [InlineData("=DOLLARFR(1.02,0)", ErrorKind.Div0)]
    public void Errors(string formula, ErrorKind expected)
    {
        Assert.Equal(Financial.Error(expected), Eval(formula));
    }

    [Fact]
    public void Built_in_functions_read_logical_values_as_numbers()
    {
        Assert.Equal(Eval("=PMT(0.1,3,100,0,1)"), Eval("=PMT(0.1,3,100,0,S!A1)"));
        Assert.Equal(Eval("=PV(0.1,3,100,0,1)"), Eval("=PV(0.1,3,100,0,TRUE)"));
    }

    [Theory]
    [InlineData("=EFFECT(0.1,S!A1)")]
    [InlineData("=NOMINAL(TRUE,2)")]
    [InlineData("=CUMIPMT(0.1,36,8000,S!A1,10,0)")]
    [InlineData("=DOLLARDE(TRUE,16)")]
    public void Analysis_toolpak_functions_refuse_logical_values(string formula)
    {
        Assert.Equal(Financial.Error(ErrorKind.Value), Eval(formula));
    }

    [Fact]
    public void Analysis_toolpak_functions_still_read_number_text()
    {
        Financial.Rounds(0.05354267, 8, Eval("=EFFECT(\"0.0525\",\"4\")"));
    }

    [Fact]
    public void Optional_arguments_left_empty_take_their_defaults()
    {
        Assert.Equal(Eval("=PMT(0.1,3,100)"), Eval("=PMT(0.1,3,100,,)"));
        Assert.Equal(Eval("=RATE(48,-200,8000)"), Eval("=RATE(48,-200,8000,,,)"));
    }

    // At 700% a period the principal paid in the first years is below 1e-20 of the payment;
    // summing principals rather than subtracting interest from the payment keeps it exact.
    [Fact]
    public void Cumulative_principal_keeps_tiny_parts()
    {
        Financial.Near(-3.308722252996902e-21, Eval("=CUMPRINC(7,36,8000,3,10,1)"));
        Financial.Near(-56000, Eval("=CUMIPMT(7,36,8000,3,10,1)"));
    }

    [Fact]
    public void Rate_solves_a_long_annuity()
    {
        Financial.Near(0.025627751816686105, Eval("=RATE(300,-200,8000,0,1)"));
    }

    [Fact]
    public void An_array_argument_gives_an_array()
    {
        _s["C1"].Formula = "=PMT({0.1;0.2},3,100)";
        _wb.Recalculate();
        Financial.Near(-40.21148036253776, _s["C1"].Value);
        Financial.Near(-47.47252747252748, _s["C2"].Value);
    }
}
