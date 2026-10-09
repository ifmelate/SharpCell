using SharpCell;

namespace SharpCell.Tests.Evaluation;

/// <summary>Behaviour the Excel corpus showed to differ from the first implementation.</summary>
public class CorpusFindingsTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public CorpusFindingsTests()
    {
        _s = _wb.AddSheet("S");
    }

    private static CellValue N(double value) => CellValue.Number(value);

    private CellValue Eval(string formula)
    {
        _s["Z100"].Formula = formula;
        _wb.Recalculate();
        return _s["Z100"].Value;
    }

    [Theory]
    [InlineData("=(-8)^(1/3)", -2.0)]
    [InlineData("=(-32)^(1/5)", -2.0)]
    [InlineData("=(-2)^3", -8.0)]
    public void Odd_roots_of_negative_numbers_are_real(string formula, double expected)
    {
        Assert.Equal(expected, Eval(formula).AsNumber(), 12);
    }

    [Theory]
    [InlineData("=(-8)^(1/2)")]
    [InlineData("=(-8)^0.4")]
    public void Other_fractional_powers_of_negative_numbers_are_NUM(string formula)
    {
        Assert.Equal(CellValue.Error(ErrorKind.Num), Eval(formula));
    }

    [Fact]
    public void INDIRECT_with_an_empty_style_argument_reads_R1C1()
    {
        _s["B2"].Value = 5;
        Assert.Equal(N(5), Eval("=INDIRECT(\"R2C2\",)"));
    }

    [Fact]
    public void OFFSET_with_a_negative_size_extends_up_and_left()
    {
        _s["A1"].Value = 1;
        _s["A2"].Value = 2;
        _s["B1"].Value = 10;
        _s["B2"].Value = 20;

        Assert.Equal(N(3), Eval("=SUM(OFFSET(A3,-1,0,-2))"));
        Assert.Equal(N(22), Eval("=SUM(OFFSET(C2,0,-1,1,-2))"));
        Assert.Equal(N(33), Eval("=SUM(OFFSET(B2,0,0,-2,-2))"));
    }

    [Fact]
    public void OFFSET_size_between_zero_and_one_counts_as_one_and_zero_is_REF()
    {
        _s["A1"].Value = 1;
        _s["B1"].Value = 10;

        Assert.Equal(N(1), Eval("=SUM(OFFSET(A1,0,0,1,0.9))"));
        Assert.Equal(N(11), Eval("=SUM(OFFSET(A1,0,0,1,2.9))"));
        Assert.Equal(CellValue.Error(ErrorKind.Ref), Eval("=SUM(OFFSET(A1,0,0,1,0))"));
        Assert.Equal(CellValue.Error(ErrorKind.Ref), Eval("=SUM(OFFSET(A1,0,0,-2,1))"));
    }

    [Theory]
    [InlineData("=OR(\"\",TRUE)", true)]
    [InlineData("=AND(\"\",TRUE)", true)]
    [InlineData("=AND(\"x\",FALSE)", false)]
    [InlineData("=OR(\"TRUE\")", true)]
    [InlineData("=AND(\"false\")", false)]
    public void AND_and_OR_skip_text_that_is_not_a_boolean(string formula, bool expected)
    {
        Assert.Equal(CellValue.Boolean(expected), Eval(formula));
    }

    [Theory]
    [InlineData("=OR(\"\")")]
    [InlineData("=AND(\"1\")")]
    [InlineData("=OR(\"\",\"\")")]
    public void AND_and_OR_with_nothing_but_skipped_text_are_VALUE(string formula)
    {
        Assert.Equal(CellValue.Error(ErrorKind.Value), Eval(formula));
    }

    [Fact]
    public void IFERROR_keeps_values_where_a_shorter_fallback_has_no_element()
    {
        _s["A1"].Formula = "=IFERROR({1;0;3;0}/{1;0;1;0},{7;8})";
        _wb.Recalculate();

        Assert.Equal([N(1), N(8), N(3), CellValue.Error(ErrorKind.NA)],
            [_s["A1"].Value, _s["A2"].Value, _s["A3"].Value, _s["A4"].Value]);
    }

    [Fact]
    public void IFERROR_broadcasts_a_column_against_a_row()
    {
        _s["A1"].Formula = "=IFERROR({1;0}/{0;0},{7,8})";
        _wb.Recalculate();

        Assert.Equal([N(7), N(8), N(7), N(8)], [_s["A1"].Value, _s["B1"].Value, _s["A2"].Value, _s["B2"].Value]);
    }

    [Fact]
    public void Optional_parameter_as_files_write_it()
    {
        _s["A1"].Value = 12;
        _s["B1"].Value = 3;

        Assert.Equal(N(6), Eval("=LAMBDA(a,_xlop.b,IF(ISOMITTED(_xlpm.b),_xlpm.a/2,_xlpm.a*_xlpm.b))(A1)"));
        Assert.Equal(N(36), Eval("=LAMBDA(a,_xlop.b,IF(ISOMITTED(_xlpm.b),_xlpm.a/2,_xlpm.a*_xlpm.b))(A1,B1)"));
        Assert.Equal("=LAMBDA(a,[b],IF(ISOMITTED(b),a/2,a*b))(A1,B1)", _s["Z100"].Formula is { } f
            ? "=" + SharpCell.Parsing.FormulaPrinter.Print(SharpCell.Parsing.FormulaParser.Parse(f, new CellAddress(100, 26)), new CellAddress(100, 26))
            : null);
    }
}
