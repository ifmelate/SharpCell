using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Evaluation;

public class LetLambdaTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public LetLambdaTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = 1;
        _s["A2"].Value = 2;
        _s["A3"].Value = 3;
    }

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    public static TheoryData<string, CellValue> Let => new()
    {
        { "=LET(x,2,x*3)", 6 },
        { "=LET(x,2,y,x+1,x*y)", 6 },
        { "=LET(x,1/0,5)", 5 },
        { "=LET(x,1,LET(x,2,x))", 2 },
        { "=LET(x,1,LET(x,2,x)+x)", 3 },
        { "=LET(r,A1:A3,SUM(r))", 6 },
        { "=LET(r,A1:A3,ROWS(r))", 3 },
        { "=let(X,2,x*3)", 6 },
        { "=LET(x,x+1,x)", Err(ErrorKind.Name) },
        { "=LET(x,y,y,1,x)", Err(ErrorKind.Name) },
        { "=LET(1,2,3)", Err(ErrorKind.Value) },
        { "=LET(x,1)", Err(ErrorKind.Value) },
        { "=LET(S!x,1,2)", Err(ErrorKind.Value) },
    };

    [Theory]
    [MemberData(nameof(Let))]
    public void LET_binds_names_lazily_and_lexically(string formula, CellValue expected) =>
        Assert.Equal(expected, _wb.Evaluate(formula));

    public static TheoryData<string, CellValue> Lambda => new()
    {
        { "=LAMBDA(x,x*2)(3)", 6 },
        { "=LAMBDA(x,y,x+y)(1,2)", 3 },
        { "=LAMBDA(7)()", 7 },
        { "=LAMBDA(x,x)(1,2)", Err(ErrorKind.Value) },
        { "=LAMBDA(x,y,x)(1)", Err(ErrorKind.Value) },
        { "=LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(1)", 1 },
        { "=LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(1,2)", 3 },
        { "=LAMBDA(x,[y],ISOMITTED(y))(1,)", true },
        { "=LAMBDA(x,ISOMITTED(x))(5)", false },
        { "=LAMBDA(x,x,x)(1,2)", Err(ErrorKind.Value) },
        { "=LAMBDA(1,x)(2)", Err(ErrorKind.Value) },
        { "=LET(f,LAMBDA(x,x+1),f(2))", 3 },
        { "=LET(k,10,f,LAMBDA(x,x+k),f(1))", 11 },
        { "=LET(k,1,f,LAMBDA(x,x+k),LET(k,100,f(0)))", 1 },
        { "=LET(add,LAMBDA(a,LAMBDA(b,a+b)),add(2)(3))", 5 },
        { "=LAMBDA(x,x+A1)(1)", 2 },
        { "=LAMBDA(r,SUM(r))(A1:A3)", 6 },
        { "=LET(x,1,x(2))", Err(ErrorKind.Value) },
        { "=(1)(2)", Err(ErrorKind.Value) },
        { "=LAMBDA(x,x)+1", Err(ErrorKind.Value) },
        { "=LAMBDA(x,x)", Err(ErrorKind.Calc) },
        { "=NOPE(1)", Err(ErrorKind.Name) },
    };

    [Theory]
    [MemberData(nameof(Lambda))]
    public void LAMBDA_calls(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    [Fact]
    public void Unused_LET_binding_is_never_evaluated()
    {
        var context = new EvaluationContext(_wb, _s, new CellAddress(1, 1));
        Evaluator.EvaluateFormula(FormulaParser.Parse("=LET(x,RAND(),1)", new CellAddress(1, 1)), context);
        Assert.False(context.UsedVolatile);
    }

    [Fact]
    public void Named_lambdas_and_recursion()
    {
        _wb.DefineName("Double", "=LAMBDA(x,x*2)");
        _wb.DefineName("Fact", "=LAMBDA(n,IF(n<=1,1,n*Fact(n-1)))");
        _wb.DefineName("Ten", "=10");
        Assert.Equal(CellValue.Number(42), _wb.Evaluate("=Double(21)"));
        Assert.Equal(CellValue.Number(120), _wb.Evaluate("=Fact(5)"));
        Assert.Equal(Err(ErrorKind.Value), _wb.Evaluate("=Ten(1)"));
    }

    [Fact]
    public void Runaway_recursion_is_NUM_error_not_a_crash()
    {
        _wb.DefineName("Down", "=LAMBDA(n,IF(n=0,0,Down(n-1)))");
        CellValue result = default;
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                result = _wb.Evaluate("=Down(100000)");
            }
            catch (Exception ex)
            {
                error = ex;
            }
        }, maxStackSize: 1024 * 1024);
        thread.Start();
        thread.Join();
        Assert.Null(error);
        Assert.Equal(Err(ErrorKind.Num), result);
    }

    [Fact]
    public void Lambda_in_a_cell_is_CALC_error()
    {
        _s["B1"].Formula = "=LAMBDA(x,x)";
        _wb.Recalculate();
        Assert.Equal(Err(ErrorKind.Calc), _s["B1"].Value);
    }

    [Fact]
    public void Dirty_lambda_argument_does_not_create_a_false_cycle()
    {
        _s["C1"].Formula = "=LAMBDA(v,IFERROR(v,C3))(C2)";
        _s["C2"].Formula = "=1";
        _s["C3"].Formula = "=C1";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(1), _s["C1"].Value);
        Assert.Empty(_wb.Diagnostics);
    }

    [Fact]
    public void Cells_read_inside_a_lambda_are_dependencies()
    {
        _s["B1"].Formula = "=LET(f,LAMBDA(x,x+A1),f(1))";
        _wb.Recalculate();
        _s["A1"].Value = 10;
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(11), _s["B1"].Value);
    }
}
