using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Functions;

/// <summary>MATCH, VLOOKUP, HLOOKUP, LOOKUP, XLOOKUP, XMATCH; expected values are Excel's.</summary>
public class LookupFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public LookupFunctionTests()
    {
        _s = _wb.AddSheet("S");
        string[] names = ["Alpha", "Beta", "Delta", "Delta", "Gamma"];
        for (var i = 0; i < names.Length; i++)
        {
            _s[i + 1, 1].Value = names[i];
            _s[i + 1, 2].Value = (i + 1) * 10;
            _s[i + 1, 3].Value = (2 * i) + 1;
        }

        // D: descending numbers. E: numbers mixed with text. F: text, "", an empty cell, 0. G: short result vector.
        _s["D1"].Value = 9;
        _s["D2"].Value = 5;
        _s["D3"].Value = 1;
        _s["E1"].Value = 1;
        _s["E2"].Value = "x";
        _s["E3"].Value = 3;
        _s["E4"].Value = "y";
        _s["E5"].Value = 5;
        _s["F1"].Value = "apple";
        _s["F2"].Value = "a*b";
        _s["F3"].Value = "";
        _s["F5"].Value = 0;
        _s["G1"].Value = "g1";
        _s["G5"].Value = "g5";
    }

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    private static CellValue Row(params CellValue[] values)
    {
        var array = new CellValue[1, values.Length];
        for (var i = 0; i < values.Length; i++)
            array[0, i] = values[i];
        return CellValue.Array(array);
    }

    private static CellValue Column(params CellValue[] values)
    {
        var array = new CellValue[values.Length, 1];
        for (var i = 0; i < values.Length; i++)
            array[i, 0] = values[i];
        return CellValue.Array(array);
    }

    public static TheoryData<string, CellValue> MatchCases => new()
    {
        { "=MATCH(\"delta\",A1:A5,0)", 3 },
        { "=MATCH(\"delta\",A1:A5)", 4 },
        { "=MATCH(\"D*\",A1:A5,0)", 3 },
        { "=MATCH(\"?ETA\",A1:A5,0)", 2 },
        { "=MATCH(\"a\",A1:A5)", Err(ErrorKind.NA) },
        { "=MATCH(6,C1:C5)", 3 },
        { "=MATCH(0,C1:C5)", Err(ErrorKind.NA) },
        { "=MATCH(100,C1:C5)", 5 },
        { "=MATCH(5,C1:C5,0)", 3 },
        { "=MATCH(4,C1:C5,0)", Err(ErrorKind.NA) },
        { "=MATCH(5,D1:D3,-1)", 2 },
        { "=MATCH(4,D1:D3,-1)", 2 },
        { "=MATCH(10,D1:D3,-1)", Err(ErrorKind.NA) },
        { "=MATCH(10,A1:B5,0)", Err(ErrorKind.NA) },
        { "=MATCH(1/0,C1:C5)", Err(ErrorKind.Div0) },
        { "=MATCH(1,C1:C5,1/0)", Err(ErrorKind.Div0) },
        { "=MATCH(\"~*\",{\"a\",\"*\"},0)", 2 },
        { "=MATCH(\"a~*b\",F1:F5,0)", 2 },
        { "=MATCH(\"\",F1:F5,0)", 3 },
        { "=MATCH(0,F1:F5,0)", 5 },
        { "=MATCH(F4,F1:F5,0)", 5 },
        { "=MATCH(TRUE,{FALSE,TRUE},0)", 2 },
        { "=MATCH(\"13\",{13},0)", Err(ErrorKind.NA) },
        { "=MATCH(13,{\"13\"},0)", Err(ErrorKind.NA) },
        { "=MATCH({5,7},C1:C5,0)", Row(3, 4) },
        { "=MATCH(9.99E+307,E1:E5)", 5 },
        { "=MATCH(\"zz\",E1:E5)", 4 },
        { "=MATCH(5,C:C,0)", 3 },
        { "=MATCH(1000,C:C)", 5 },
    };

    [Theory]
    [MemberData(nameof(MatchCases))]
    public void MATCH(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    public static TheoryData<string, CellValue> TableLookupCases => new()
    {
        { "=VLOOKUP(\"beta\",A1:B5,2,FALSE)", 20 },
        { "=VLOOKUP(\"Delta\",A1:B5,2)", 40 },
        { "=VLOOKUP(\"Delta\",A1:B5,2,FALSE)", 30 },
        { "=VLOOKUP(\"Zeta\",A1:B5,2)", 50 },
        { "=VLOOKUP(\"A\",A1:B5,2)", Err(ErrorKind.NA) },
        { "=VLOOKUP(\"G*\",A1:B5,2,FALSE)", 50 },
        { "=VLOOKUP(\"Alpha\",A1:B5,3)", Err(ErrorKind.Ref) },
        { "=VLOOKUP(\"Nope\",A1:B5,3,FALSE)", Err(ErrorKind.Ref) },
        { "=VLOOKUP(\"Alpha\",A1:B5,0)", Err(ErrorKind.Value) },
        { "=VLOOKUP(\"Alpha\",A1:B5,1/0)", Err(ErrorKind.Div0) },
        { "=VLOOKUP(1/0,A1:B5,2)", Err(ErrorKind.Div0) },
        { "=VLOOKUP(\"x\",\"text\",1)", Err(ErrorKind.Value) },
        { "=VLOOKUP(\"x\",3,1)", Err(ErrorKind.NA) },
        { "=VLOOKUP(2,{1,\"a\";2,\"b\"},2,FALSE)", "b" },
        { "=VLOOKUP(\"Alpha\",A1:B5,{1,2},FALSE)", Row("Alpha", 10) },
        { "=VLOOKUP(6,C:C,1)", 5 },
        { "=HLOOKUP(2,{1,2,3;\"a\",\"b\",\"c\"},2,FALSE)", "b" },
        { "=HLOOKUP(2.5,{1,2,3;\"a\",\"b\",\"c\"},2)", "b" },
        { "=HLOOKUP(0,{1,2,3;\"a\",\"b\",\"c\"},2)", Err(ErrorKind.NA) },
        { "=HLOOKUP(1,{1,2},3)", Err(ErrorKind.Ref) },
        { "=HLOOKUP(\"Alpha\",A1:B5,1,FALSE)", "Alpha" },
    };

    [Theory]
    [MemberData(nameof(TableLookupCases))]
    public void VLOOKUP_and_HLOOKUP(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    public static TheoryData<string, CellValue> LookupCases => new()
    {
        { "=LOOKUP(9.99E+307,E1:E5)", 5 },
        { "=LOOKUP(\"zz\",E1:E5)", "y" },
        { "=LOOKUP(2,1/(C1:C5>4),B1:B5)", 50 },
        { "=LOOKUP(6,C1:C5,B1:B5)", 30 },
        { "=LOOKUP(0,C1:C5,B1:B5)", Err(ErrorKind.NA) },
        { "=LOOKUP(\"b\",{\"a\",\"b\",\"c\";1,2,3})", 2 },
        { "=LOOKUP(2,{1,\"a\";2,\"b\";3,\"c\"})", "b" },
        { "=LOOKUP(4,{1;3;5},{\"a\";\"b\";\"c\"})", "b" },
        { "=LOOKUP(4,{1;3;5},{\"a\",\"b\",\"c\"})", "b" },
        { "=LOOKUP(3,{1,2,3,4,5})", 3 },
        { "=LOOKUP(9,C1:C5,G1)", "g5" },
        { "=LOOKUP(9,C1:C5,{1,2})", Err(ErrorKind.NA) },
        { "=LOOKUP(9,C1:C5,A1:B2)", Err(ErrorKind.NA) },
    };

    [Theory]
    [MemberData(nameof(LookupCases))]
    public void LOOKUP(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    public static TheoryData<string, CellValue> XLookupCases => new()
    {
        { "=XLOOKUP(\"Delta\",A1:A5,B1:B5)", 30 },
        { "=XLOOKUP(\"Delta\",A1:A5,B1:B5,,,-1)", 40 },
        { "=XLOOKUP(\"nope\",A1:A5,B1:B5)", Err(ErrorKind.NA) },
        { "=XLOOKUP(\"nope\",A1:A5,B1:B5,\"none\")", "none" },
        { "=XLOOKUP(\"nope\",A1:A5,B1:B5,{1,2})", Row(1, 2) },
        { "=XLOOKUP(6,C1:C5,B1:B5,,-1)", 30 },
        { "=XLOOKUP(6,C1:C5,B1:B5,,1)", 40 },
        { "=XLOOKUP(0,C1:C5,B1:B5,,-1)", Err(ErrorKind.NA) },
        { "=XLOOKUP(\"g*\",A1:A5,B1:B5,,2)", 50 },
        { "=XLOOKUP(\"g*\",A1:A5,B1:B5)", Err(ErrorKind.NA) },
        { "=XLOOKUP(\"g*\",A1:A5,B1:B5,,2,2)", Err(ErrorKind.Value) },
        { "=XLOOKUP(1,A1:A5,B1:B4)", Err(ErrorKind.Value) },
        { "=XLOOKUP(1,A1:B5,B1:B5)", Err(ErrorKind.Value) },
        { "=XLOOKUP(1,C1:C5,B1:B5,,5)", Err(ErrorKind.Value) },
        { "=XLOOKUP(1,C1:C5,B1:B5,,0,3)", Err(ErrorKind.Value) },
        { "=XLOOKUP(1,C1:C5,B1:B5,,,TRUE)", 10 },
        { "=XLOOKUP(1/0,C1:C5,B1:B5)", Err(ErrorKind.Div0) },
        { "=XLOOKUP(3,C1:C5,A1:B5)", Row("Beta", 20) },
        { "=XLOOKUP(10,A1:B1,A2:B3)", Column(20, 30) },
        { "=SUM(XLOOKUP(3,C1:C5,B1:B5):XLOOKUP(7,C1:C5,B1:B5))", 90 },
        { "=XLOOKUP(5,{1,5,5,5,9},{1,2,3,4,5},,0,2)", 2 },
        { "=XLOOKUP(5,{9,5,5,5,1},{1,2,3,4,5},,0,-2)", 4 },
        { "=XLOOKUP(6,{9,5,5,5,1},{1,2,3,4,5},,-1,-2)", 2 },
        { "=XLOOKUP(6,{9,5,5,5,1},{1,2,3,4,5},,1,-2)", 1 },
        { "=XLOOKUP(6,{1,5,5,5,9},{1,2,3,4,5},,-1,2)", 4 },
        { "=XLOOKUP(6,{1,5,5,5,9},{1,2,3,4,5},,1,2)", 5 },
        { "=XLOOKUP(\"b\",{\"a\",\"b\"},{1,2},,1,2)", 2 },
        { "=XLOOKUP(\"z\",{1,2,\"a\"},{1,2,3},,-1)", 3 },
        { "=XLOOKUP(\"z\",{1,2,TRUE},{1,2,3},,1)", 3 },
        { "=XLOOKUP(\"^D\",A1:A5,B1:B5,,3)", 30 },
        { "=XLOOKUP(\"(\",A1:A5,B1:B5,,3)", Err(ErrorKind.Value) },
        { "=XLOOKUP(H1,F1:F5,B1:B5)", 40 },
        { "=XLOOKUP(,F1:F5,B1:B5,,,-1)", 40 },
        { "=XLOOKUP(TRUE,{1,\"a\",TRUE},{1,2,3})", 3 },
        { "=XLOOKUP(1,{1},{7,8})", Row(7, 8) },
        { "=XLOOKUP({3;5},C1:C5,B1:B5)", Column(20, 30) },
        { "=XLOOKUP(7,C:C,B:B)", 40 },
        { "=XLOOKUP(7,C:C,B:B,,,2)", 40 },
    };

    [Theory]
    [MemberData(nameof(XLookupCases))]
    public void XLOOKUP(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    public static TheoryData<string, CellValue> XMatchCases => new()
    {
        { "=XMATCH(\"Delta\",A1:A5)", 3 },
        { "=XMATCH(\"Delta\",A1:A5,0,-1)", 4 },
        { "=XMATCH(6,C1:C5,1)", 4 },
        { "=XMATCH(6,C1:C5,-1)", 3 },
        { "=XMATCH(4,C1:C5,0,2)", Err(ErrorKind.NA) },
        { "=XMATCH(7,C1:C5,0,2)", 4 },
        { "=XMATCH(7,C1:C5,0,-2)", Err(ErrorKind.NA) },
        { "=XMATCH(1,A1:B2)", Err(ErrorKind.Value) },
        { "=XMATCH(\"?amma\",A1:A5,2)", 5 },
        { "=XMATCH(\"a.*a\",A1:A5,3)", 5 },
        { "=XMATCH(42,A1:A5,3)", Err(ErrorKind.NA) },
    };

    [Theory]
    [MemberData(nameof(XMatchCases))]
    public void XMATCH(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    [Fact]
    public void Lookups_follow_their_inputs()
    {
        _s["H1"].Formula = "=XLOOKUP(\"Gamma\",A:A,B:B)";
        _s["H2"].Formula = "=VLOOKUP(\"Gamma\",A1:B5,2,FALSE)";
        _s["H3"].Formula = "=LOOKUP(\"Gamma\",A:A,B:B)";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(50), _s["H1"].Value);

        _s["B5"].Value = 55;
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(55), _s["H1"].Value);
        Assert.Equal(CellValue.Number(55), _s["H2"].Value);
        Assert.Equal(CellValue.Number(55), _s["H3"].Value);
    }

    [Fact]
    public void Lookups_over_whole_columns_walk_only_stored_cells()
    {
        _s["C1000000"].Value = 99;
        _s["B1000000"].Value = "last";
        Assert.Equal(CellValue.Text("last"), _wb.Evaluate("=XLOOKUP(99,C:C,B:B)"));

        var before = GC.GetAllocatedBytesForCurrentThread();
        Assert.Equal(CellValue.Text("last"), _wb.Evaluate("=XLOOKUP(99,C:C,B:B,,,-1)"));
        Assert.Equal(CellValue.Number(1000000), _wb.Evaluate("=MATCH(99,C:C,0)"));
        Assert.Equal(CellValue.Number(1000000), _wb.Evaluate("=MATCH(1000,C:C)"));
        Assert.Equal(CellValue.Number(99), _wb.Evaluate("=VLOOKUP(99,C:C,1)"));
        Assert.True(GC.GetAllocatedBytesForCurrentThread() - before < 4_000_000);
    }

    [Fact]
    public void LOOKUP_vectors_are_array_parameters_in_legacy_formulas()
    {
        var loader = new SheetLoader(_s);
        loader.SetFormula(2, 9, FormulaParser.Parse("=LOOKUP(2,1/(C1:C5>4),B1:B5)", new CellAddress(2, 9)), LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.SetFormula(3, 9, FormulaParser.Parse("=MATCH(TRUE,C1:C5>4,0)", new CellAddress(3, 9)), LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.Complete();
        _wb.Recalculate();

        Assert.Equal(CellValue.Number(50), _s["I2"].Value);

        // MATCH's array is not an array parameter: in old Excel this needed Ctrl+Shift+Enter, and
        // without it the comparison is intersected with the formula's row (C3>4 is TRUE).
        Assert.Equal(CellValue.Number(1), _s["I3"].Value);
    }
}
