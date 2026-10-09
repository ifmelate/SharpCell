using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Functions;

public class LogicalFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public LogicalFunctionTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = 0;
        _s["A2"].Value = 1;
        _s["A3"].Value = 0;
        _s["B1"].Value = 10;
        _s["B2"].Value = 20;
        _s["B3"].Value = 30;
        _s["C1"].Value = "Monday";
        _s["C2"].Value = CellValue.Error(ErrorKind.NA);
        _s["C3"].Value = CellValue.Error(ErrorKind.Div0);
    }

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    private static CellValue Arr(CellValue[,] values) => CellValue.Array(values);

    public static TheoryData<string, CellValue> Cases => new()
    {
        { "=TRUE()", true },
        { "=FALSE()", false },
        { "=IFNA(C2,\"none\")", "none" },
        { "=IFNA(C3,\"none\")", Err(ErrorKind.Div0) },
        { "=IFNA(5,1/0)", 5 },
        { "=IFNA(NA(),)", 0 },
        { "=IFNA(Z99,1)", 0 },
        { "=IFNA({1,#N/A,#DIV/0!},0)", Arr(new CellValue[,] { { 1, 0, Err(ErrorKind.Div0) } }) },
        { "=IFS(FALSE,1,TRUE,2)", 2 },
        { "=IFS(0,1,5,2)", 2 },
        { "=IFS(FALSE,1)", Err(ErrorKind.NA) },
        { "=IFS(Z99,1)", Err(ErrorKind.NA) },
        { "=IFS(\"1\",1)", Err(ErrorKind.Value) },
        { "=IFS(\"true\",1)", 1 },
        { "=IFS(C3,1,TRUE,2)", Err(ErrorKind.Div0) },
        { "=IFS(TRUE,1/0,TRUE,2)", Err(ErrorKind.Div0) },
        { "=IFS(TRUE,1,C3,2)", 1 },
        { "=IFS(TRUE,Z99)", 0 },
        { "=IFS({TRUE,FALSE,FALSE},1,{FALSE,TRUE,FALSE},2)", Arr(new CellValue[,] { { 1, 2, Err(ErrorKind.NA) } }) },
        { "=IFS(FALSE,1,{TRUE,FALSE},{5,6},TRUE,9)", Arr(new CellValue[,] { { 5, 9 } }) },
        { "=SWITCH(2,1,\"a\",2,\"b\")", "b" },
        { "=SWITCH(3,1,\"a\",2,\"b\")", Err(ErrorKind.NA) },
        { "=SWITCH(3,1,\"a\",2,\"b\",\"other\")", "other" },
        { "=SWITCH(C1,\"MONDAY\",1,\"Tuesday\",2)", 1 },
        { "=SWITCH(C3,1,\"a\",\"b\")", Err(ErrorKind.Div0) },
        { "=SWITCH(1,1/0,\"a\",\"b\")", Err(ErrorKind.Div0) },
        { "=SWITCH(1,1,\"a\",1/0,\"b\")", "a" },
        { "=SWITCH(1,\"1\",\"text\",TRUE,\"bool\",\"none\")", "none" },
        { "=SWITCH(1,1,1/0)", Err(ErrorKind.Div0) },
        { "=SWITCH({1,2,3},1,\"a\",2,\"b\",\"c\")", Arr(new CellValue[,] { { "a", "b", "c" } }) },
        { "=XOR(TRUE)", true },
        { "=XOR(TRUE,TRUE)", false },
        { "=XOR(TRUE,TRUE,TRUE)", true },
        { "=XOR(1,0)", true },
        { "=XOR(A1:A3)", true },
        { "=XOR(A1:A3,B1)", false },
        { "=XOR(\"TRUE\")", true },
        { "=XOR(\"\",FALSE)", false },
        { "=XOR(\"\")", Err(ErrorKind.Value) },
        { "=XOR(\"1\")", Err(ErrorKind.Value) },
        { "=XOR(Z99)", Err(ErrorKind.Value) },
        { "=XOR(C1)", Err(ErrorKind.Value) },
        { "=XOR(TRUE,C3)", Err(ErrorKind.Div0) },
        { "=XOR({TRUE,FALSE,TRUE})", false },
    };

    [Theory]
    [MemberData(nameof(Cases))]
    public void Logical_functions(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    [Theory]
    [InlineData("=IFNA(B1,C1)")]
    [InlineData("=IFS(ISERROR(B1),C1,TRUE,B1)")]
    [InlineData("=SWITCH(ISERROR(B1),TRUE,C1,B1)")]
    public void Dirty_selector_does_not_create_a_false_cycle(string formula)
    {
        var wb = new Workbook();
        var s = wb.AddSheet("S");
        s["A1"].Formula = formula;
        s["B1"].Formula = "=1";
        s["C1"].Formula = "=A1";
        wb.Recalculate();

        Assert.Equal(CellValue.Number(1), s["A1"].Value);
        Assert.Equal(CellValue.Number(1), s["C1"].Value);
        Assert.Empty(wb.Diagnostics);
    }

    [Fact]
    public void Unchosen_branches_are_not_evaluated()
    {
        // A loop through an unchosen branch is no loop.
        _s["D1"].Formula = "=IFS(TRUE,1,TRUE,D1)";
        _s["D2"].Formula = "=SWITCH(1,1,2,D2)";
        _s["D3"].Formula = "=IFNA(1,D3)";
        _wb.Recalculate();

        Assert.Equal(CellValue.Number(1), _s["D1"].Value);
        Assert.Equal(CellValue.Number(2), _s["D2"].Value);
        Assert.Equal(CellValue.Number(1), _s["D3"].Value);
        Assert.Empty(_wb.Diagnostics);
    }

    private CellValue Legacy(string formula, int row = 2, int column = 5)
    {
        var loader = new SheetLoader(_s);
        loader.SetFormula(row, column, FormulaParser.Parse(formula, new CellAddress(row, column)), LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.Complete();
        _wb.Recalculate();
        return _s[row, column].Value;
    }

    [Fact]
    public void Legacy_formulas_intersect_conditions_and_values()
    {
        // Row 2: A2 is 1, B2 is 20, C2 is #N/A.
        Assert.Equal(CellValue.Text("y"), Legacy("=IFS(A1:A3,\"y\",TRUE,\"n\")"));
        Assert.Equal(CellValue.Text("b"), Legacy("=SWITCH(B1:B3,10,\"a\",20,\"b\")"));
        Assert.Equal(CellValue.Number(7), Legacy("=IFNA(C1:C3,7)"));
    }
}
