namespace SharpCell.Tests.Evaluation;

public class FunctionValueTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public FunctionValueTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = 1;
        _s["B1"].Value = 2;
        _s["A2"].Value = 3;
        _s["B2"].Value = 4;
        _s["D1"].Value = "ab";
        _s["D2"].Value = "cde";
    }

    [Theory]
    [InlineData("=BYROW(S!A1:B2,SUM)", "{3;7}")]
    [InlineData("=BYCOL(S!A1:B2,MAX)", "{3,4}")]
    [InlineData("=MAP(S!D1:D2,LEN)", "{2;3}")]
    [InlineData("=MAP(S!A1:A2,S!B1:B2,SUM)", "{3;7}")]
    [InlineData("=REDUCE(0,S!A1:B2,SUM)", "10")]
    [InlineData("=SCAN(0,S!A1:A2,SUM)", "{1;4}")]
    [InlineData("=MAKEARRAY(2,2,SUM)", "{2,3;3,4}")]
    [InlineData("=LET(f,SUM,f(1,2,3))", "6")]
    [InlineData("=LET(f,len,f(\"abcd\"))", "4")]
    [InlineData("=MAP(S!A1:A2,ABS)", "{1;3}")]
    [InlineData("=BYROW(S!A1:B2,MAX)", "{2;4}")]
    public void A_function_name_is_a_callable_value(string formula, string expected)
    {
        Assert.Equal(expected, _wb.Evaluate(formula).ToString());
    }

    [Fact]
    public void Files_write_a_function_value_with_the_eta_prefix()
    {
        _s["F1"].Formula = "=_xlfn.BYROW(S!A1:B2,_xleta.SUM)";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(7), _s["F2"].Value);
        Assert.Equal("=_xlfn.BYROW(S!A1:B2,_xleta.SUM)", _s["F1"].Formula);
    }

    [Fact]
    public void Scalar_parameters_apply_element_by_element_through_the_value()
    {
        Assert.Equal("{1,#VALUE!,1}", _wb.Evaluate("=MAP({-1,\"x\",TRUE},ABS)").ToString());
    }

    [Fact]
    public void A_defined_name_shadows_a_built_in_function_value()
    {
        _wb.DefineName("LEN", "=LAMBDA(x, 99)");
        Assert.Equal("{99;99}", _wb.Evaluate("=MAP(S!D1:D2,LEN)").ToString());
    }

    [Fact]
    public void A_custom_function_is_a_value_too_and_its_callers_follow_it()
    {
        _s["F1"].Formula = "=MAP(S!A1:A2,TWICE)";
        _wb.Recalculate();
        Assert.Equal(CellValue.Error(ErrorKind.Name), _s["F1"].Value);

        _wb.Functions.Add("TWICE", a => a.Number(0).AsNumber() * 2);
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(6), _s["F2"].Value);
    }

    [Theory]
    [InlineData("=BYROW(S!A1:B2,NOSUCHFN)")]
    [InlineData("=LET(f,LET,f(1))")]
    [InlineData("=LET(f,LAMBDA,f(1))")]
    [InlineData("=BYROW(S!A1:B2,WEBSERVICE)")]
    public void Names_that_are_not_callable_functions_stay_name_errors(string formula)
    {
        Assert.Equal(CellValue.Error(ErrorKind.Name), _wb.Evaluate(formula));
    }

    [Theory]
    [InlineData("=LET(f,ABS,f(1,2))")]
    [InlineData("=LET(f,ABS,f())")]
    public void A_wrong_number_of_arguments_is_a_value_error(string formula)
    {
        Assert.Equal(CellValue.Error(ErrorKind.Value), _wb.Evaluate(formula));
    }

    [Fact]
    public void A_volatile_function_value_makes_its_caller_volatile()
    {
        _s["G1"].Formula = "=LET(f,RAND,f())";
        _wb.Recalculate();
        var before = _wb.Calculation.EvaluationCount;
        _wb.Recalculate();
        Assert.True(_wb.Calculation.EvaluationCount > before);
    }
}
