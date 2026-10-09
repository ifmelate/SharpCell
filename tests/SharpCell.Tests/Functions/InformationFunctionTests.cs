using SharpCell;

namespace SharpCell.Tests.Functions;

public class InformationFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public InformationFunctionTests()
    {
        _s = _wb.AddSheet("S");
        _wb.AddSheet("Other");
        _wb.AddSheet("Third sheet");
        _s["A1"].Value = 3.5;
        _s["A2"].Value = "text";
        _s["A3"].Value = true;
        _s["A4"].Value = CellValue.Error(ErrorKind.NA);
        _s["A5"].Value = CellValue.Error(ErrorKind.Div0);
        _s["A6"].Value = "12";
        _s["B1"].Formula = "=1+1";
        _s["B2"].Formula = "=\"x\"";
        _wb.DefineName("Seven", "=7");
        _wb.DefineName("Cells", "=S!$A$1:$A$3");
        _wb.Recalculate();
    }

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    private static CellValue Arr(CellValue[,] values) => CellValue.Array(values);

    public static TheoryData<string, CellValue> Predicates => new()
    {
        { "=ISERR(A5)", true },
        { "=ISERR(A4)", false },
        { "=ISERR(1)", false },
        { "=ISNA(A4)", true },
        { "=ISNA(A5)", false },
        { "=ISNA(NA())", true },
        { "=ISLOGICAL(A3)", true },
        { "=ISLOGICAL(1)", false },
        { "=ISLOGICAL(\"TRUE\")", false },
        { "=ISNUMBER(A1)", true },
        { "=ISNUMBER(A6)", false },
        { "=ISNUMBER(Z99)", false },
        { "=ISTEXT(A2)", true },
        { "=ISTEXT(Z99)", false },
        { "=ISNONTEXT(Z99)", true },
        { "=ISNONTEXT(A2)", false },
        { "=ISNONTEXT(A4)", true },
        { "=ISNUMBER(A1:A3)", Arr(new CellValue[,] { { true }, { false }, { false } }) },
        { "=ISEVEN(2)", true },
        { "=ISEVEN(2.9)", true },
        { "=ISEVEN(-4.5)", true },
        { "=ISODD(-5.5)", true },
        { "=ISODD(A6)", false },
        { "=ISEVEN(Z99)", true },
        { "=ISODD(A3)", Err(ErrorKind.Value) },
        { "=ISODD(TRUE)", Err(ErrorKind.Value) },
        { "=ISODD(A2)", Err(ErrorKind.Value) },
        { "=ISEVEN(A4)", Err(ErrorKind.NA) },
        { "=ISREF(A1)", true },
        { "=ISREF(A1:B2)", true },
        { "=ISREF((A1,B2))", true },
        { "=ISREF(Other!A1)", true },
        { "=ISREF(INDIRECT(\"A1\"))", true },
        { "=ISREF(Cells)", true },
        { "=ISREF(Seven)", false },
        { "=ISREF(\"A1\")", false },
        { "=ISREF({1,2})", false },
        { "=ISREF(NoSuchName)", false },
        { "=ISREF(1/0)", false },
        { "=LET(x,A1,ISREF(x))", true },
        { "=ISFORMULA(B1)", true },
        { "=ISFORMULA(A1)", false },
        { "=ISFORMULA(Z99)", false },
        { "=ISFORMULA(A1:B1)", Arr(new CellValue[,] { { false, true } }) },
        { "=ISFORMULA(5)", Err(ErrorKind.Value) },
        { "=ISFORMULA(1/0)", Err(ErrorKind.Div0) },
    };

    [Theory]
    [MemberData(nameof(Predicates))]
    public void IS_functions(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    public static TheoryData<string, CellValue> Values => new()
    {
        { "=N(A1)", 3.5 },
        { "=N(A2)", 0 },
        { "=N(A3)", 1 },
        { "=N(A6)", 0 },
        { "=N(Z99)", 0 },
        { "=N(A5)", Err(ErrorKind.Div0) },
        { "=N(TRUE+TRUE)", 2 },
        { "=NA()", Err(ErrorKind.NA) },
        { "=TYPE(A1)", 1 },
        { "=TYPE(Z99)", 1 },
        { "=TYPE(A2)", 2 },
        { "=TYPE(A3)", 4 },
        { "=TYPE(A4)", 16 },
        { "=TYPE(A1:A3)", 64 },
        { "=TYPE({1,2})", 64 },
        { "=TYPE(LAMBDA(x,x))", 128 },
        { "=ERROR.TYPE(#NULL!)", 1 },
        { "=ERROR.TYPE(A5)", 2 },
        { "=ERROR.TYPE(#VALUE!)", 3 },
        { "=ERROR.TYPE(#REF!)", 4 },
        { "=ERROR.TYPE(#NAME?)", 5 },
        { "=ERROR.TYPE(#NUM!)", 6 },
        { "=ERROR.TYPE(A4)", 7 },
        { "=ERROR.TYPE(#SPILL!)", 9 },
        { "=ERROR.TYPE(#CALC!)", 14 },
        { "=ERROR.TYPE(1)", Err(ErrorKind.NA) },
        { "=ERROR.TYPE(Z99)", Err(ErrorKind.NA) },
        { "=SHEET()", 1 },
        { "=SHEET(Other!A1)", 2 },
        { "=SHEET(\"third SHEET\")", 3 },
        { "=SHEET(\"Missing\")", Err(ErrorKind.NA) },
        { "=SHEET(\"\")", Err(ErrorKind.NA) },
        { "=SHEET(0)", Err(ErrorKind.NA) },
        { "=SHEET(1/0)", Err(ErrorKind.Div0) },
        { "=SHEETS()", 3 },
        { "=SHEETS(A1:B5)", 1 },
        { "=SHEETS('S:Third sheet'!A1)", 3 },
        { "=SHEETS(\"S\")", Err(ErrorKind.Value) },
        { "=SHEETS(1/0)", Err(ErrorKind.Div0) },
    };

    [Theory]
    [MemberData(nameof(Values))]
    public void Value_functions(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    public static TheoryData<string, CellValue> CellCases => new()
    {
        { "=CELL(\"address\",B3)", "$B$3" },
        { "=CELL(\"address\",C2:D4)", "$C$2" },
        { "=CELL(\"address\",Other!AA10)", "Other!$AA$10" },
        { "=CELL(\"col\",C2:D4)", 3 },
        { "=CELL(\"row\",C2:D4)", 2 },
        { "=CELL(\"contents\",A1:A3)", 3.5 },
        { "=CELL(\"contents\",B1)", 2 },
        { "=CELL(\"contents\",A5)", Err(ErrorKind.Div0) },
        { "=CELL(\"contents\",Z99)", 0 },
        { "=CELL(\"type\",Z99)", "b" },
        { "=CELL(\"type\",A1)", "v" },
        { "=CELL(\"Type\",A2)", "l" },
        { "=CELL(\"type\",A3)", "v" },
        { "=CELL(\"type\",A4)", "v" },
        { "=CELL(\"type\",B2)", "v" },
        { "=CELL(\"format\",A1)", Err(ErrorKind.NA) },
        { "=CELL(\"hello\",A1)", Err(ErrorKind.Value) },
        { "=CELL(\"\",A1)", Err(ErrorKind.Value) },
        { "=CELL(TRUE,A1)", Err(ErrorKind.Value) },
        { "=CELL(1/0,A1)", Err(ErrorKind.Div0) },
        { "=CELL(\"type\",INDIRECT(\"A\"&10^9))", Err(ErrorKind.Ref) },
        { "=CELL(\"type\",5)", Err(ErrorKind.Value) },
        { "=INFO(\"numfile\")", 3 },
        { "=INFO(\"RECALC\")", "Automatic" },
        { "=INFO(\"directory\")", Err(ErrorKind.NA) },
        { "=INFO(\"nonsense\")", Err(ErrorKind.Value) },
    };

    [Theory]
    [MemberData(nameof(CellCases))]
    public void CELL_and_INFO(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    [Fact]
    public void CELL_without_a_reference_describes_its_own_cell()
    {
        _s["D7"].Formula = "=CELL(\"address\")";
        _wb.Recalculate();
        Assert.Equal(CellValue.Text("$D$7"), _s["D7"].Value);
    }

    [Fact]
    public void ISFORMULA_of_its_own_cell_is_no_loop()
    {
        _s["C9"].Formula = "=ISFORMULA(INDIRECT(\"C9\"))";
        _wb.Recalculate();
        Assert.Equal(CellValue.True, _s["C9"].Value);
        Assert.Empty(_wb.Diagnostics);
    }

    [Fact]
    public void Spilled_cells_count_as_formulas()
    {
        _s["E1"].Formula = "={1;2}";
        _s["F1"].Formula = "=ISFORMULA(E2)";
        _wb.Recalculate();
        Assert.Equal(CellValue.True, _s["F1"].Value);
    }

    [Fact]
    public void Results_follow_edits_of_the_cells_they_describe()
    {
        _s["G1"].Formula = "=ISFORMULA(H1)";
        _s["G2"].Formula = "=CELL(\"type\",H1)";
        _s["G3"].Formula = "=CELL(\"contents\",H1)";
        _s["H1"].Value = "abc";
        _wb.Recalculate();
        Assert.Equal([CellValue.False, CellValue.Text("l"), CellValue.Text("abc")], [_s["G1"].Value, _s["G2"].Value, _s["G3"].Value]);

        _s["H1"].Formula = "=2*3";
        _wb.Recalculate();
        Assert.Equal([CellValue.True, CellValue.Text("v"), CellValue.Number(6)], [_s["G1"].Value, _s["G2"].Value, _s["G3"].Value]);
    }

    [Fact]
    public void SHEETS_counts_sheets_added_later()
    {
        _s["G5"].Formula = "=SHEETS()";
        _wb.Recalculate();
        _wb.AddSheet("Fourth");
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(4), _s["G5"].Value);
    }
}
