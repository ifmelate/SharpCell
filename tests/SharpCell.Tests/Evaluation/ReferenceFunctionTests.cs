using SharpCell;

namespace SharpCell.Tests.Evaluation;

public class ReferenceFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public ReferenceFunctionTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = 1;
        _s["A2"].Value = 2;
        _s["A3"].Value = 3;
        _s["B2"].Value = 20;
        _wb.DefineName("Rate", "=S!$B$2");
    }

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    public static TheoryData<string, CellValue> Cases => new()
    {
        { "=OFFSET(A1,1,1)", 20 },
        { "=SUM(OFFSET(A1,0,0,3,1))", 6 },
        { "=ROWS(OFFSET(A1,0,0,5,2))", 5 },
        { "=COLUMNS(OFFSET(A1:B3,1,1))", 2 },
        { "=SUM(OFFSET(A1,1,0,,))", 2 },
        { "=OFFSET(A1,-1,0)", Err(ErrorKind.Ref) },
        { "=OFFSET(A1,0,0,0,1)", Err(ErrorKind.Ref) },
        { "=OFFSET(XFD1,0,1)", Err(ErrorKind.Ref) },
        { "=OFFSET(5,1,1)", Err(ErrorKind.Value) },
        { "=OFFSET(A1,\"x\",0)", Err(ErrorKind.Value) },
        { "=SUM(OFFSET(A1,{0,1},0))", 3 },
        { "=INDIRECT(\"A2\")", 2 },
        { "=INDIRECT(\"b2\")", 20 },
        { "=SUM(INDIRECT(\"S!A1:A3\"))", 6 },
        { "=SUM(INDIRECT(\"'S'!A1:S!A2\"))", 3 },
        { "=INDIRECT(\"R2C1\",FALSE)", 2 },
        { "=INDIRECT(\"R[1]C\",FALSE)", 2 },
        { "=INDIRECT(\"Rate\")", 20 },
        { "=SUM(INDIRECT({\"A1\",\"A2\"}))", 3 },
        { "=INDIRECT(\"nope!A1\")", Err(ErrorKind.Ref) },
        { "=INDIRECT(\"1+1\")", Err(ErrorKind.Ref) },
        { "=INDIRECT(\"\")", Err(ErrorKind.Ref) },
        { "=INDIRECT(\"=A1\")", Err(ErrorKind.Ref) },
        { "=INDIRECT(\"A1:\")", Err(ErrorKind.Ref) },
        { "=LET(x,5,INDIRECT(\"x\"))", Err(ErrorKind.Ref) },
    };

    [Theory]
    [MemberData(nameof(Cases))]
    public void OFFSET_and_INDIRECT(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    [Fact]
    public void INDIRECT_follows_its_target_and_its_inputs()
    {
        _s["D2"].Value = 1;
        _s["D1"].Formula = "=INDIRECT(\"A\"&D2)";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(1), _s["D1"].Value);

        _s["D2"].Value = 3;
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(3), _s["D1"].Value);

        _s["A3"].Value = 30;
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(30), _s["D1"].Value);
    }

    [Fact]
    public void OFFSET_and_INDIRECT_are_volatile()
    {
        _s["D1"].Formula = "=OFFSET(A1,0,0)";
        _s["D2"].Formula = "=INDIRECT(\"A1\")";
        _wb.Recalculate();
        var before = _wb.Calculation.EvaluationCount;
        _wb.Recalculate();
        Assert.Equal(2, _wb.Calculation.EvaluationCount - before);
    }

    [Fact]
    public void Spill_of_an_OFFSET_range()
    {
        _s["F1"].Formula = "=OFFSET(A1,0,0,3,1)*10";
        _wb.Recalculate();
        Assert.Equal([CellValue.Number(10), CellValue.Number(20), CellValue.Number(30)], [_s["F1"].Value, _s["F2"].Value, _s["F3"].Value]);
    }
}
