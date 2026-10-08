using SharpCell;

namespace SharpCell.Tests.Evaluation;

public class LambdaFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public LambdaFunctionTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = 1;
        _s["A2"].Value = 2;
        _s["A3"].Value = 3;
        _s["B1"].Value = 10;
        _s["B2"].Value = 20;
    }

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    private static CellValue Arr(CellValue[,] values) => CellValue.Array(values);

    public static TheoryData<string, CellValue> Cases => new()
    {
        { "=MAP({1,2,3},LAMBDA(x,x*2))", Arr(new CellValue[,] { { 2, 4, 6 } }) },
        { "=MAP({1,2},{10,20},LAMBDA(a,b,a+b))", Arr(new CellValue[,] { { 11, 22 } }) },
        { "=MAP(A1:A3,LAMBDA(x,x+1))", Arr(new CellValue[,] { { 2 }, { 3 }, { 4 } }) },
        { "=MAP({1,2},LAMBDA(a,b,a))", Err(ErrorKind.Value) },
        { "=MAP({1},5)", Err(ErrorKind.Value) },
        { "=MAP({1,2},LAMBDA(x,{1,2}))", Arr(new CellValue[,] { { Err(ErrorKind.Calc), Err(ErrorKind.Calc) } }) },
        { "=REDUCE(0,{1,2,3},LAMBDA(a,v,a+v))", 6 },
        { "=REDUCE(,A1:A3,LAMBDA(a,v,a+v))", 6 },
        { "=REDUCE(1,{2,3},LAMBDA(a,v,a*v))", 6 },
        { "=REDUCE(0,{1,2},LAMBDA(a,v,{1,2}*v))", Arr(new CellValue[,] { { 2, 4 } }) },
        { "=SCAN(0,{1,2,3},LAMBDA(a,v,a+v))", Arr(new CellValue[,] { { 1, 3, 6 } }) },
        { "=SCAN(0,{1,2;3,4},LAMBDA(a,v,a+v))", Arr(new CellValue[,] { { 1, 3 }, { 6, 10 } }) },
        { "=BYROW({1,2;3,4},LAMBDA(r,SUM(r)))", Arr(new CellValue[,] { { 3 }, { 7 } }) },
        { "=BYROW(A1:B2,LAMBDA(r,SUM(r)))", Arr(new CellValue[,] { { 11 }, { 22 } }) },
        { "=BYROW(A1:B2,LAMBDA(r,COLUMNS(r)))", Arr(new CellValue[,] { { 2 }, { 2 } }) },
        { "=BYROW({1,2;3,4},LAMBDA(r,r))", Arr(new CellValue[,] { { Err(ErrorKind.Calc) }, { Err(ErrorKind.Calc) } }) },
        { "=BYCOL({1,2;3,4},LAMBDA(c,SUM(c)))", Arr(new CellValue[,] { { 4, 6 } }) },
        { "=BYCOL(A1:B2,LAMBDA(c,SUM(c)))", Arr(new CellValue[,] { { 3, 30 } }) },
        { "=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))", Arr(new CellValue[,] { { 11, 12, 13 }, { 21, 22, 23 } }) },
        { "=MAKEARRAY(0,1,LAMBDA(r,c,1))", Err(ErrorKind.Value) },
        { "=MAKEARRAY(100000,1000,LAMBDA(r,c,1))", Err(ErrorKind.Num) },
        { "=MAKEARRAY(1,1,LAMBDA(r,c,r+c))", Arr(new CellValue[,] { { 2 } }) },
    };

    [Theory]
    [MemberData(nameof(Cases))]
    public void Lambda_helper_functions(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    [Fact]
    public void Lambdas_are_not_called_on_dirty_inputs()
    {
        _s["D1"].Formula = "=SUM(MAP(E1:E3,LAMBDA(v,IFERROR(v,D5))))";
        _s["E1"].Formula = "=1";
        _s["E2"].Formula = "=2";
        _s["E3"].Formula = "=3";
        _s["D5"].Formula = "=D1";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(6), _s["D1"].Value);
        Assert.Empty(_wb.Diagnostics);
    }

    [Fact]
    public void Results_spill()
    {
        _s["G1"].Formula = "=MAKEARRAY(2,2,LAMBDA(r,c,r+c))";
        _wb.Recalculate();
        Assert.Equal([CellValue.Number(2), CellValue.Number(3), CellValue.Number(3), CellValue.Number(4)],
            [_s["G1"].Value, _s["H1"].Value, _s["G2"].Value, _s["H2"].Value]);
    }
}
