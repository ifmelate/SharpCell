using SharpCell;

namespace SharpCell.Tests.Functions;

public class TextSplitFunctionTests
{
    private readonly Workbook _wb = new();

    public TextSplitFunctionTests()
    {
        var s = _wb.AddSheet("S");
        s["A1"].Value = ",";
        s["B1"].Value = ".";
        s["C1"].Value = "-";
        s["A2"].Value = "123 the 456 the 891";
        s["A3"].Value = CellValue.Error(ErrorKind.Div0);
    }

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    private static CellValue Arr(CellValue[,] values) => CellValue.Array(values);

    private void Check(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    public static TheoryData<string, CellValue> Splits => new()
    {
        { "=TEXTSPLIT(\"a,b,c\",\",\")", Arr(new CellValue[,] { { "a", "b", "c" } }) },
        { "=TEXTSPLIT(\"1,2;3,4\",\",\",\";\")", Arr(new CellValue[,] { { "1", "2" }, { "3", "4" } }) },
        { "=TEXTSPLIT(\"1,2,3;4\",\",\",\";\")", Arr(new CellValue[,] { { "1", "2", "3" }, { "4", Err(ErrorKind.NA), Err(ErrorKind.NA) } }) },
        { "=TEXTSPLIT(\"1,2,3;4\",\",\",\";\",,,\"-\")", Arr(new CellValue[,] { { "1", "2", "3" }, { "4", "-", "-" } }) },
        { "=TEXTSPLIT(\"1,,2\",\",\")", Arr(new CellValue[,] { { "1", "", "2" } }) },
        { "=TEXTSPLIT(\"1,,2\",\",\",,TRUE)", Arr(new CellValue[,] { { "1", "2" } }) },
        { "=TEXTSPLIT(\"1,2.3-4\",{\",\",\".\",\"-\"})", Arr(new CellValue[,] { { "1", "2", "3", "4" } }) },
        { "=TEXTSPLIT(\"1,2.3-4\",A1:C1)", Arr(new CellValue[,] { { "1", "2", "3", "4" } }) },
        { "=TEXTSPLIT(\"1,2.3-4\",,{\",\",\".\"})", Arr(new CellValue[,] { { "1" }, { "2" }, { "3-4" } }) },
        { "=TEXTSPLIT(\"aXbxc\",\"x\")", Arr(new CellValue[,] { { "aXb", "c" } }) },
        { "=TEXTSPLIT(\"aXbxc\",\"x\",,,1)", Arr(new CellValue[,] { { "a", "b", "c" } }) },
        { "=TEXTSPLIT(\"abc\",\",\")", "abc" },
        { "=TEXTSPLIT(\"abc\",\"\")", Err(ErrorKind.Value) },
        { "=TEXTSPLIT(\"abc\",)", Err(ErrorKind.Value) },
        { "=TEXTSPLIT(\"abc\",,\"\")", Err(ErrorKind.Value) },
        { "=TEXTSPLIT(\"abc\",\",\",,,2)", Err(ErrorKind.Value) },
        { "=TEXTSPLIT(A3,\",\")", Err(ErrorKind.Div0) },
        { "=TEXTSPLIT(\"a,b\",A3)", Err(ErrorKind.Div0) },
    };

    [Theory]
    [MemberData(nameof(Splits))]
    public void Textsplit(string formula, CellValue expected) => Check(formula, expected);

    public static TheoryData<string, CellValue> BeforeAfter => new()
    {
        { "=TEXTBEFORE(A2,\"the\")", "123 " },
        { "=TEXTAFTER(A2,\"the\")", " 456 the 891" },
        { "=TEXTBEFORE(A2,\"the\",2)", "123 the 456 " },
        { "=TEXTAFTER(A2,\"the\",-1)", " 891" },
        { "=TEXTAFTER(A2,\"the\",-2)", " 456 the 891" },
        { "=TEXTAFTER(A2,\"the\",1.6)", " 456 the 891" },
        { "=TEXTAFTER(A2,\"the\",-1.2)", " 456 the 891" },
        { "=TEXTAFTER(A2,\"THE\")", Err(ErrorKind.NA) },
        { "=TEXTAFTER(A2,\"THE\",1,1)", " 456 the 891" },
        { "=TEXTAFTER(A2,\"the\",3)", Err(ErrorKind.NA) },
        { "=TEXTAFTER(A2,\"the\",3,,,\"none\")", "none" },
        { "=TEXTAFTER(A2,\"the\",0)", Err(ErrorKind.Value) },
        { "=TEXTAFTER(A2,\"the\",30)", Err(ErrorKind.Value) },
        { "=TEXTAFTER(\"del\",\"deli\")", Err(ErrorKind.NA) },
        { "=TEXTAFTER(\"del\",\"deli\",1)", Err(ErrorKind.Value) },
        { "=TEXTAFTER(A2,\"x\",1,,1)", "" },
        { "=TEXTBEFORE(A2,\"x\",1,,1)", "123 the 456 the 891" },
        { "=TEXTAFTER(A2,\"x\",-1,,1)", "123 the 456 the 891" },
        { "=TEXTBEFORE(A2,\"x\",-1,,1)", "" },
        { "=TEXTAFTER(A2,\"x\",2,,1)", Err(ErrorKind.NA) },
        { "=TEXTAFTER(A2,\"x\",1,,2)", Err(ErrorKind.Value) },
        { "=TEXTAFTER(A2,\"x\",1,3)", Err(ErrorKind.Value) },
        { "=TEXTBEFORE(\"abc\",\"\")", "" },
        { "=TEXTAFTER(\"abc\",\"\")", "abc" },
        { "=TEXTBEFORE(\"abc\",\"\",-1)", "abc" },
        { "=TEXTAFTER(\"abc\",\"\",-1)", "" },
        { "=TEXTAFTER(\"a-b.c\",{\".\",\"-\"})", "b.c" },
        { "=TEXTAFTER(TRUE,\"r\",1,1)", "UE" },
        { "=TEXTAFTER(123,1)", "23" },
        { "=TEXTAFTER(A3,\"x\")", Err(ErrorKind.Div0) },
        { "=TEXTAFTER(A2,\"the\",A3)", Err(ErrorKind.Div0) },
        { "=TEXTAFTER(A2,\"x\",,,,A3)", Err(ErrorKind.Div0) },
        { "=TEXTAFTER(A2,\"the\",,,,A3)", " 456 the 891" },
        { "=TEXTAFTER(A2,\"x\",,,,Z9)", "" },
        { "=TEXTBEFORE({\"a-b\",\"c-d\"},\"-\")", Arr(new CellValue[,] { { "a", "c" } }) },
    };

    [Theory]
    [MemberData(nameof(BeforeAfter))]
    public void Textbefore_and_textafter(string formula, CellValue expected) => Check(formula, expected);
}
