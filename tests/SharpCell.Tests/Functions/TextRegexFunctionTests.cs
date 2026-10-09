using SharpCell;

namespace SharpCell.Tests.Functions;

public class TextRegexFunctionTests
{
    private readonly Workbook _wb = new();

    public TextRegexFunctionTests() => _wb.AddSheet("S");

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    private static CellValue Row(params CellValue[] values)
    {
        var array = new CellValue[1, values.Length];
        for (var i = 0; i < values.Length; i++)
            array[0, i] = values[i];
        return CellValue.Array(array);
    }

    public static TheoryData<string, CellValue> Cases => new()
    {
        { "=REGEXTEST(\"Mr. Sherlock Holmes\",\"[A-Z]her\")", true },
        { "=REGEXTEST(\"Mr. Sherlock Holmes\",\"[0-9]+\")", false },
        { "=REGEXTEST(\"abc\",\"B\")", false },
        { "=REGEXTEST(\"abc\",\"B\",1)", true },
        { "=REGEXTEST(\"abc\",\"B\",2)", Err(ErrorKind.Value) },
        { "=REGEXTEST(\"abc\",\"(\")", Err(ErrorKind.Value) },
        { "=REGEXTEST(1/0,\"a\")", Err(ErrorKind.Div0) },
        { "=REGEXTEST({\"a1\",\"b\"},\"\\d\")", Row(true, false) },
        { "=REGEXEXTRACT(\"LonelyPlanet\",\"[A-Z][a-z]+\")", "Lonely" },
        { "=REGEXEXTRACT(\"LonelyPlanet\",\"[A-Z][a-z]+\",1)", Row("Lonely", "Planet") },
        { "=REGEXEXTRACT(\"John Smith\",\"(\\w+) (\\w+)\",2)", Row("John", "Smith") },
        { "=REGEXEXTRACT(\"abc\",\"\\d\")", Err(ErrorKind.NA) },
        { "=REGEXEXTRACT(\"abc\",\"\\d\",1)", Err(ErrorKind.NA) },
        { "=REGEXEXTRACT(\"abc\",\"b\",3)", Err(ErrorKind.Value) },
        { "=REGEXEXTRACT(\"ABC\",\"b\",0,1)", "B" },
        { "=REGEXREPLACE(\"JamesBond\",\"([A-Z][a-z]+)([A-Z][a-z]+)\",\"$2, $1 $2\")", "Bond, James Bond" },
        { "=REGEXREPLACE(\"a1b2c3\",\"\\d\",\"#\")", "a#b#c#" },
        { "=REGEXREPLACE(\"a1b2c3\",\"\\d\",\"#\",2)", "a1b#c3" },
        { "=REGEXREPLACE(\"a1b2c3\",\"\\d\",\"#\",-1)", "a1b2c#" },
        { "=REGEXREPLACE(\"a1b2c3\",\"\\d\",\"#\",4)", "a1b2c3" },
        { "=REGEXREPLACE(\"aBc\",\"b\",\"x\",0,1)", "axc" },
        { "=REGEXREPLACE(\"abc\",\"[\",\"x\")", Err(ErrorKind.Value) },
    };

    [Theory]
    [MemberData(nameof(Cases))]
    public void Regular_expressions(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    [Fact]
    public void A_runaway_pattern_times_out_as_an_error() =>
        Assert.Equal(Err(ErrorKind.Value), _wb.Evaluate("=REGEXTEST(REPT(\"a\",5000)&\"!\",\"^(a|a?)+$\")"));
}
