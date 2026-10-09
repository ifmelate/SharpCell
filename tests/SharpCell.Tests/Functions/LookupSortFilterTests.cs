using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Functions;

/// <summary>SORT, SORTBY, FILTER, UNIQUE; expected values are Excel's. Arrays are written as formula literals.</summary>
public class LookupSortFilterTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public LookupSortFilterTests()
    {
        // A1:B5: names and scores; A6:B6 empty; D1:D4 mixed types with a gap at D3.
        _s = _wb.AddSheet("S");
        string[] names = ["Hola", "de", "HOLA", "Ana", "beta"];
        double[] scores = [3, 1, 2, 3, 1];
        for (var i = 0; i < names.Length; i++)
        {
            _s[i + 1, 1].Value = names[i];
            _s[i + 1, 2].Value = scores[i];
        }

        _s["D1"].Value = true;
        _s["D2"].Value = "b";
        _s["D4"].Value = 5;
        _s["D5"].Formula = "=1/0";
        _s["D6"].Value = -1;
        _wb.Recalculate();
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    private CellValue Literal(string array) => _wb.Evaluate("=" + array);

    public static TheoryData<string, string> Cases => new()
    {
        { "=SORT({3;1;2})", "{1;2;3}" },
        { "=SORT({3;1;2},,-1)", "{3;2;1}" },
        { "=SORT(A1:A5)", "{\"Ana\";\"beta\";\"de\";\"Hola\";\"HOLA\"}" },
        { "=SORT(A1:A5,1,-1)", "{\"Hola\";\"HOLA\";\"de\";\"beta\";\"Ana\"}" },
        { "=SORT(A1:B5,2)", "{\"de\",1;\"beta\",1;\"HOLA\",2;\"Hola\",3;\"Ana\",3}" },
        { "=SORT(A1:B5,{2,1},{-1,1})", "{\"Ana\",3;\"Hola\",3;\"HOLA\",2;\"beta\",1;\"de\",1}" },
        { "=SORT({3,1,2;30,10,20},1,1,TRUE)", "{1,2,3;10,20,30}" },
        { "=SORT({3,1,2;30,10,20},2,-1,TRUE)", "{3,2,1;30,20,10}" },
        { "=SORTBY(A1:A5,B1:B5)", "{\"de\";\"beta\";\"HOLA\";\"Hola\";\"Ana\"}" },
        { "=SORTBY(A1:A5,B1:B5,-1,A1:A5,1)", "{\"Ana\";\"Hola\";\"HOLA\";\"beta\";\"de\"}" },
        { "=SORTBY({\"a\",\"b\",\"c\"},{3,1,2})", "{\"b\",\"c\",\"a\"}" },
        { "=FILTER(A1:B5,B1:B5=3)", "{\"Hola\",3;\"Ana\",3}" },
        { "=FILTER(A1:A5,(B1:B5=1)+(A1:A5=\"Ana\"))", "{\"de\";\"Ana\";\"beta\"}" },
        { "=FILTER({1,2,3;4,5,6},{1,0,1})", "{1,3;4,6}" },
        { "=FILTER(A1:A5,B1:B5>5,\"none\")", "\"none\"" },
        { "=FILTER(A1:A5,B1:B5>5,{1,2})", "{1,2}" },
        { "=UNIQUE(A1:A5)", "{\"Hola\";\"de\";\"Ana\";\"beta\"}" },
        { "=UNIQUE(B1:B5)", "{3;1;2}" },
        { "=UNIQUE(B1:B5,,TRUE)", "{2}" },
        { "=UNIQUE(A1:B5)", "{\"Hola\",3;\"de\",1;\"HOLA\",2;\"Ana\",3;\"beta\",1}" },
        { "=UNIQUE({1,2,1;3,4,3},TRUE)", "{1,2;3,4}" },
    };

    [Fact]
    public void Types_sort_numbers_text_booleans_errors_with_empty_cells_last()
    {
        var empty = CellValue.Empty;
        var div0 = CellValue.Error(ErrorKind.Div0);
        Assert.Equal(CellValue.Array(new CellValue[,] { { -1 }, { 5 }, { "b" }, { true }, { div0 }, { empty } }), Eval("=SORT(D1:D6)"));
        Assert.Equal(CellValue.Array(new CellValue[,] { { div0 }, { true }, { "b" }, { 5 }, { -1 }, { empty } }), Eval("=SORT(D1:D6,1,-1)"));
        Assert.Equal(CellValue.Array(new CellValue[,] { { "Ana" }, { "beta" }, { empty } }), Eval("=UNIQUE(A4:A8)"));
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void Sorting_and_filtering(string formula, string expected) => Assert.Equal(Literal(expected), Eval(formula));

    public static TheoryData<string, ErrorKind> Errors => new()
    {
        { "=SORT(A1:B5,3)", ErrorKind.Value },
        { "=SORT(A1:B5,0)", ErrorKind.Value },
        { "=SORT(A1:B5,1,2)", ErrorKind.Value },
        { "=SORT(A1:B5,{1,2},{1,1,1})", ErrorKind.Value },
        { "=SORT(1/0)", ErrorKind.Div0 },
        { "=SORTBY(A1:A5,B1:B4)", ErrorKind.Value },
        { "=SORTBY(A1:A5,B1:B5,0)", ErrorKind.Value },
        { "=SORTBY(A1:B5,B1:B5,1,{1,2},1)", ErrorKind.Value },
        { "=FILTER(A1:A5,B1:B5>5)", ErrorKind.Calc },
        { "=FILTER(A1:A5,A1:A5)", ErrorKind.Value },
        { "=FILTER(A1:A5,B1:B4>1)", ErrorKind.Value },
        { "=FILTER(D4:D6,D4:D6)", ErrorKind.Div0 },
        { "=UNIQUE(B1:B5*0,,TRUE)", ErrorKind.Calc },
    };

    [Theory]
    [MemberData(nameof(Errors))]
    public void Sorting_and_filtering_errors(string formula, ErrorKind expected) => Assert.Equal(CellValue.Error(expected), Eval(formula));

    [Fact]
    public void Results_spill_and_follow_their_inputs()
    {
        _s["F1"].Formula = "=SORT(FILTER(B1:B5,B1:B5>1),,-1)";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(3), _s["F1"].Value);
        Assert.Equal(CellValue.Number(3), _s["F2"].Value);
        Assert.Equal(CellValue.Number(2), _s["F3"].Value);
        Assert.Equal(CellValue.Empty, _s["F4"].Value);

        _s["B2"].Value = 9;
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(9), _s["F1"].Value);
        Assert.Equal(CellValue.Number(2), _s["F4"].Value);
    }

    [Fact]
    public void Array_inputs_are_array_parameters_in_legacy_formulas()
    {
        var loader = new SheetLoader(_s);
        loader.SetFormula(1, 8, FormulaParser.Parse("=SUM(TRANSPOSE(B1:B5*2))", new CellAddress(1, 8)), LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.Complete();
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(20), _s["H1"].Value);
    }
}
