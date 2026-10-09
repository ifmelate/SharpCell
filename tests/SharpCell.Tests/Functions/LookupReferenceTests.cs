using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Functions;

/// <summary>INDEX, ROW, COLUMN, ADDRESS, AREAS, FORMULATEXT, TRIMRANGE; expected values are Excel's.</summary>
public class LookupReferenceTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public LookupReferenceTests()
    {
        _s = _wb.AddSheet("S");
        string[] names = ["Alpha", "Beta", "Delta", "Delta", "Gamma"];
        for (var i = 0; i < names.Length; i++)
        {
            _s[i + 1, 1].Value = names[i];
            _s[i + 1, 2].Value = (i + 1) * 10;
            _s[i + 1, 3].Value = (2 * i) + 1;
        }

        // T: a sparse block at K3:L5 for TRIMRANGE.
        var t = _wb.AddSheet("T");
        t["K3"].Value = 1;
        t["L5"].Value = "x";
    }

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    private static CellValue Array(CellValue[,] values) => CellValue.Array(values);

    public static TheoryData<string, CellValue> IndexCases => new()
    {
        { "=INDEX({1,2,3},2)", 2 },
        { "=INDEX({1;2;3},2)", 2 },
        { "=INDEX({1,2;3,4},2,1)", 3 },
        { "=INDEX({1,2;3,4},0,2)", Array(new CellValue[,] { { 2 }, { 4 } }) },
        { "=INDEX({1,2;3,4},1)", Array(new CellValue[,] { { 1, 2 } }) },
        { "=INDEX({1,2;3,4},,)", Array(new CellValue[,] { { 1, 2 }, { 3, 4 } }) },
        { "=INDEX({1,2,3},1,0,1)", Array(new CellValue[,] { { 1, 2, 3 } }) },
        { "=INDEX({1,2},-1)", Err(ErrorKind.Value) },
        { "=INDEX({1,2},3)", Err(ErrorKind.Ref) },
        { "=INDEX({1,2;3,4},1,3)", Err(ErrorKind.Ref) },
        { "=INDEX({1,2},1,1,2)", Err(ErrorKind.Ref) },
        { "=INDEX({1,2},\"x\")", Err(ErrorKind.Value) },
        { "=INDEX(1/0,1)", Err(ErrorKind.Div0) },
        { "=INDEX(7,1)", 7 },
        { "=INDEX(B1:B5,3)", 30 },
        { "=INDEX(A1:C1,2)", 10 },
        { "=INDEX(B1:C5,2,2)", 3 },
        { "=SUM(INDEX(B1:C5,0,2))", 25 },
        { "=SUM(INDEX(B1:C5,2,0))", 23 },
        { "=SUM(B1:INDEX(B1:B5,3))", 60 },
        { "=ROWS(INDEX(B1:C5,0,1))", 5 },
        { "=INDEX((A1:A5,B1:B5),2,1,2)", 20 },
        { "=INDEX((A1:A5,B1:B5),1,1,3)", Err(ErrorKind.Ref) },
        { "=INDEX(B1:B5,{1;3})", Array(new CellValue[,] { { 10 }, { 30 } }) },
        { "=INDEX(B:B,MATCH(\"Gamma\",A:A,0))", 50 },
        { "=INDEX(B1:B5*2,3)", 60 },
    };

    [Theory]
    [MemberData(nameof(IndexCases))]
    public void INDEX(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    public static TheoryData<string, CellValue> PositionCases => new()
    {
        { "=ROW(C3)", 3 },
        { "=COLUMN(C3)", 3 },
        { "=ROW($E$7)", 7 },
        { "=ROW(B2:C4)", Array(new CellValue[,] { { 2 }, { 3 }, { 4 } }) },
        { "=COLUMN(B2:D3)", Array(new CellValue[,] { { 2, 3, 4 } }) },
        { "=SUM(ROW(A1:A3))", 6 },
        { "=SUM(ROW(A1:A3)*COLUMN(A1:C1))", 36 },
        { "=MAX(ROW(A:A))", 1048576 },
        { "=MAX(COLUMN(1:1))", 16384 },
        { "=ROW(OFFSET(A1,2,3))", 3 },
        { "=ROW({1,2})", Err(ErrorKind.Value) },
        { "=ROW(1/0)", Err(ErrorKind.Div0) },
        { "=ROW((A1,B2))", Err(ErrorKind.Ref) },
        { "=AREAS(A1:B2)", 1 },
        { "=AREAS((A1,B2,C3))", 3 },
        { "=AREAS(1)", Err(ErrorKind.Value) },
    };

    [Theory]
    [MemberData(nameof(PositionCases))]
    public void ROW_COLUMN_and_AREAS(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    [Fact]
    public void ROW_and_COLUMN_without_argument_are_the_formula_cell()
    {
        _s["H2"].Formula = "=ROW()";
        _s["H3"].Formula = "=COLUMN()";
        _s["XFD9"].Formula = "=COLUMN()*1000+ROW()";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(2), _s["H2"].Value);
        Assert.Equal(CellValue.Number(8), _s["H3"].Value);
        Assert.Equal(CellValue.Number(16384009), _s["XFD9"].Value);
    }

    public static TheoryData<string, CellValue> AddressCases => new()
    {
        { "=ADDRESS(2,3)", "$C$2" },
        { "=ADDRESS(2,3,2)", "C$2" },
        { "=ADDRESS(2,3,3)", "$C2" },
        { "=ADDRESS(2,3,4)", "C2" },
        { "=ADDRESS(2,3,1,FALSE)", "R2C3" },
        { "=ADDRESS(2,3,2,FALSE)", "R2C[3]" },
        { "=ADDRESS(2,3,3,FALSE)", "R[2]C3" },
        { "=ADDRESS(2,3,4,FALSE)", "R[2]C[3]" },
        { "=ADDRESS(23.7,34.2,2)", "AH$23" },
        { "=ADDRESS(1048576,16384)", "$XFD$1048576" },
        { "=ADDRESS(2,3,1,TRUE,\"Sheet2\")", "Sheet2!$C$2" },
        { "=ADDRESS(1,1,1,TRUE,\"My Sheet\")", "'My Sheet'!$A$1" },
        { "=ADDRESS(1,1,1,TRUE,\"It's\")", "'It''s'!$A$1" },
        { "=ADDRESS(1,1,1,TRUE,\"A1\")", "'A1'!$A$1" },
        { "=ADDRESS(1,1,1,TRUE,\"2024\")", "'2024'!$A$1" },
        { "=ADDRESS(0,1)", Err(ErrorKind.Value) },
        { "=ADDRESS(-1,1)", Err(ErrorKind.Value) },
        { "=ADDRESS(1,16385)", Err(ErrorKind.Value) },
        { "=ADDRESS(1,1,5)", Err(ErrorKind.Value) },
        { "=ADDRESS(1,1,0)", Err(ErrorKind.Value) },
        { "=ADDRESS(1/0,1)", Err(ErrorKind.Div0) },
        { "=INDIRECT(ADDRESS(2,2))", 20 },
    };

    [Theory]
    [MemberData(nameof(AddressCases))]
    public void ADDRESS(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    [Fact]
    public void FORMULATEXT_reads_the_formula_and_follows_it()
    {
        _s["H1"].Formula = "=SUM(B1:B5)";
        _s["H2"].Formula = "=FORMULATEXT(H1)";
        _s["H3"].Formula = "=FORMULATEXT(H3)";
        _s["H4"].Formula = "=FORMULATEXT(A1)";
        _s["H5"].Formula = "=FORMULATEXT(1)";
        _s["H6"].Formula = "=FORMULATEXT(H1:H2)";
        _wb.Recalculate();
        Assert.Equal(CellValue.Text("=SUM(B1:B5)"), _s["H2"].Value);
        Assert.Equal(CellValue.Text("=FORMULATEXT(H3)"), _s["H3"].Value);
        Assert.Equal(Err(ErrorKind.NA), _s["H4"].Value);
        Assert.Equal(Err(ErrorKind.Value), _s["H5"].Value);
        Assert.Equal(CellValue.Text("=SUM(B1:B5)"), _s["H6"].Value);

        _s["H1"].Formula = "=MAX(B1:B5)";
        _wb.Recalculate();
        Assert.Equal(CellValue.Text("=MAX(B1:B5)"), _s["H2"].Value);
    }

    public static TheoryData<string, CellValue> TrimRangeCases => new()
    {
        { "=ROWS(TRIMRANGE(T!A1:Z20))", 3 },
        { "=COLUMNS(TRIMRANGE(T!A1:Z20))", 2 },
        { "=ROW(TRIMRANGE(T!A1:Z20))", Array(new CellValue[,] { { 3 }, { 4 }, { 5 } }) },
        { "=ROWS(TRIMRANGE(T!A1:Z20,1))", 18 },
        { "=ROWS(TRIMRANGE(T!A1:Z20,2))", 5 },
        { "=ROWS(TRIMRANGE(T!A1:Z20,0))", 20 },
        { "=COLUMNS(TRIMRANGE(T!A1:Z20,3,0))", 26 },
        { "=COLUMNS(TRIMRANGE(T!A1:Z20,3,1))", 16 },
        { "=COLUMNS(TRIMRANGE(T!A1:Z20,3,2))", 12 },
        { "=ROWS(TRIMRANGE(T!K:K))", 1 },
        { "=SUM(TRIMRANGE(T!A:Z))", 1 },
        { "=TRIMRANGE(T!K3:L5)", Array(new CellValue[,] { { 1, CellValue.Empty }, { CellValue.Empty, CellValue.Empty }, { CellValue.Empty, "x" } }) },
        { "=TRIMRANGE(T!A1:Z20,4)", Err(ErrorKind.Value) },
        { "=TRIMRANGE({1,2})", Err(ErrorKind.Value) },
        { "=INDIRECT(\"R1C1:R2\",FALSE)", Err(ErrorKind.Ref) },
    };

    [Theory]
    [MemberData(nameof(TrimRangeCases))]
    public void TRIMRANGE(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    [Fact]
    public void TRIMRANGE_grows_with_its_range()
    {
        _s["H1"].Formula = "=ROWS(TRIMRANGE(A:A))";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(5), _s["H1"].Value);

        _s["A9"].Value = "x";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(9), _s["H1"].Value);
    }

    [Fact]
    public void INDEX_array_is_an_array_parameter_in_legacy_formulas()
    {
        var loader = new SheetLoader(_s);
        loader.SetFormula(2, 9, FormulaParser.Parse("=INDEX(B1:B5*2,4)", new CellAddress(2, 9)), LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.Complete();
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(80), _s["I2"].Value);
    }

    [Fact]
    public void INDEX_result_follows_its_inputs()
    {
        _s["H1"].Formula = "=INDEX(B:B,MATCH(\"Gamma\",A:A,0))";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(50), _s["H1"].Value);

        _s["B5"].Value = 55;
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(55), _s["H1"].Value);
    }
}
