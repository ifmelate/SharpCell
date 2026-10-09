using SharpCell;

namespace SharpCell.Tests.Functions;

/// <summary>
/// TRANSPOSE, TAKE, DROP, CHOOSEROWS, CHOOSECOLS, EXPAND, HSTACK, VSTACK, TOCOL, TOROW, WRAPROWS,
/// WRAPCOLS; expected values are Excel's. Arrays are written as formula literals.
/// </summary>
public class LookupArrayShapeTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public LookupArrayShapeTests()
    {
        // A1:C3 holds 1..9 row by row; B5 is the only stored cell of A4:C6 besides C6.
        _s = _wb.AddSheet("S");
        for (var r = 1; r <= 3; r++)
        {
            for (var c = 1; c <= 3; c++)
                _s[r, c].Value = ((r - 1) * 3) + c;
        }

        _s["B5"].Value = "x";
        _s["C6"].Formula = "=1/0";
        _wb.Recalculate();
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    private CellValue Literal(string array) => _wb.Evaluate("=" + array);

    public static TheoryData<string, string> Cases => new()
    {
        { "=TRANSPOSE(A1:C2)", "{1,4;2,5;3,6}" },
        { "=TRANSPOSE({1,2,3})", "{1;2;3}" },
        { "=TRANSPOSE(5)", "{5}" },
        { "=TAKE(A1:C3,2)", "{1,2,3;4,5,6}" },
        { "=TAKE(A1:C3,-1)", "{7,8,9}" },
        { "=TAKE(A1:C3,,2)", "{1,2;4,5;7,8}" },
        { "=TAKE(A1:C3,2,-1)", "{3;6}" },
        { "=TAKE(A1:C3,100)", "{1,2,3;4,5,6;7,8,9}" },
        { "=TAKE(A1:C3,1.9)", "{1,2,3}" },
        { "=TAKE(A:C,1)", "{1,2,3}" },
        { "=DROP(A1:C3,2)", "{7,8,9}" },
        { "=DROP(A1:C3,-2)", "{1,2,3}" },
        { "=DROP(A1:C3,1,1)", "{5,6;8,9}" },
        { "=DROP(A1:C3,,-2)", "{1;4;7}" },
        { "=CHOOSEROWS(A1:C3,3,1,1)", "{7,8,9;1,2,3;1,2,3}" },
        { "=CHOOSEROWS(A1:C3,-1)", "{7,8,9}" },
        { "=CHOOSEROWS(A1:C3,{2,3})", "{4,5,6;7,8,9}" },
        { "=CHOOSECOLS(A1:C3,3,-3)", "{3,1;6,4;9,7}" },
        { "=EXPAND({1,2},2)", "{1,2;#N/A,#N/A}" },
        { "=EXPAND({1,2},2,3,0)", "{1,2,0;0,0,0}" },
        { "=EXPAND({1,2},,3,\"-\")", "{1,2,\"-\"}" },
        { "=HSTACK({1;2},{3})", "{1,3;2,#N/A}" },
        { "=HSTACK(1,\"a\",TRUE)", "{1,\"a\",TRUE}" },
        { "=HSTACK(A1:A2,C1:C2)", "{1,3;4,6}" },
        { "=VSTACK({1,2},{3})", "{1,2;3,#N/A}" },
        { "=VSTACK(A1:B1,A3:B3)", "{1,2;7,8}" },
        { "=TOCOL(A1:B2)", "{1;2;4;5}" },
        { "=TOCOL(A1:B2,0,TRUE)", "{1;4;2;5}" },
        { "=TOROW(A1:B2)", "{1,2,4,5}" },
        { "=TOROW(A1:B2,,TRUE)", "{1,4,2,5}" },
        { "=TOCOL(A4:C6,1)", "{\"x\";#DIV/0!}" },
        { "=TOCOL(A4:C6,3)", "{\"x\"}" },
        { "=TOCOL({1,#N/A;3,4},2,TRUE)", "{1;3;4}" },
        { "=TOCOL(A:C,1,TRUE)", "{1;4;7;2;5;8;\"x\";3;6;9;#DIV/0!}" },
        { "=WRAPROWS({1,2,3,4,5},2)", "{1,2;3,4;5,#N/A}" },
        { "=WRAPROWS({1;2;3},2,0)", "{1,2;3,0}" },
        { "=WRAPROWS({1,2,3},5)", "{1,2,3,#N/A,#N/A}" },
        { "=WRAPCOLS({1,2,3,4,5},2)", "{1,3,5;2,4,#N/A}" },
        { "=WRAPCOLS(FALSE,2)", "{FALSE;#N/A}" },
    };

    [Theory]
    [MemberData(nameof(Cases))]
    public void Reshaping(string formula, string expected) => Assert.Equal(Literal(expected), Eval(formula));

    public static TheoryData<string, ErrorKind> Errors => new()
    {
        { "=TAKE(A1:C3,0)", ErrorKind.Calc },
        { "=TAKE(A1:C3,1,0)", ErrorKind.Calc },
        { "=DROP(A1:C3,3)", ErrorKind.Calc },
        { "=DROP(A1:C3,-7)", ErrorKind.Calc },
        { "=TAKE(A1:C3,\"x\")", ErrorKind.Value },
        { "=TAKE(1/0,1)", ErrorKind.Div0 },
        { "=CHOOSEROWS(A1:C3,0)", ErrorKind.Value },
        { "=CHOOSEROWS(A1:C3,4)", ErrorKind.Value },
        { "=CHOOSECOLS(A1:C3,-4)", ErrorKind.Value },
        { "=CHOOSECOLS(A1:C3,1/0)", ErrorKind.Div0 },
        { "=EXPAND(A1:C3,2)", ErrorKind.Value },
        { "=EXPAND(A1:C3,3,2)", ErrorKind.Value },
        { "=EXPAND(1,2000000)", ErrorKind.Num },
        { "=VSTACK(1,(A1,B2))", ErrorKind.Value },
        { "=TOCOL(A1:B2,4)", ErrorKind.Value },
        { "=TOCOL(A4:A6,1)", ErrorKind.Calc },
        { "=WRAPROWS(A1:C3,2)", ErrorKind.Value },
        { "=WRAPROWS({1,2},0)", ErrorKind.Num },
    };

    [Theory]
    [MemberData(nameof(Errors))]
    public void Reshaping_errors(string formula, ErrorKind expected) => Assert.Equal(CellValue.Error(expected), Eval(formula));

    [Fact]
    public void Empty_cells_spill_as_zero()
    {
        _s["E1"].Formula = "=TRANSPOSE(A4:C4)";
        _s["G1"].Formula = "=TAKE(A4:C5,-1)";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(0), _s["E1"].Value);
        Assert.Equal(CellValue.Number(0), _s["E3"].Value);
        Assert.Equal(CellValue.Number(0), _s["G1"].Value);
        Assert.Equal(CellValue.Text("x"), _s["H1"].Value);
    }

    [Fact]
    public void Shapes_follow_their_inputs()
    {
        _s["E1"].Formula = "=SUM(TOCOL(A:C,3))";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(45), _s["E1"].Value);

        _s["A2"].Value = 104;
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(145), _s["E1"].Value);
    }
}
