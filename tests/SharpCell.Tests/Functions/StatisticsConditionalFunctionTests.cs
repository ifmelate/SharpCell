using SharpCell;

namespace SharpCell.Tests.Functions;

public class StatisticsConditionalFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public StatisticsConditionalFunctionTests()
    {
        _s = _wb.AddSheet("S");
        string[] names = ["a", "b", "a", "c", "a", "b"];
        object[] values = [10, 20, 40, "x", true, -5];
        for (var i = 0; i < names.Length; i++)
        {
            _s[$"A{i + 1}"].Value = names[i];
            _s[$"B{i + 1}"].Value = values[i] switch
            {
                int n => n,
                string t => t,
                bool b => b,
                _ => CellValue.Empty,
            };
        }

        _s["C1"].Value = "e";
        _s["D1"].Formula = "=1/0";
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    [Theory]
    [InlineData("=AVERAGEIF(A1:A6,\"a\",B1:B6)", 25)]
    [InlineData("=AVERAGEIF(A1:A6,\"a\",B1)", 25)]
    [InlineData("=AVERAGEIF(B1:B6,\">0\")", 70.0 / 3)]
    [InlineData("=AVERAGEIF(A:A,\"b\",B:B)", 7.5)]
    [InlineData("=AVERAGEIFS(B1:B6,A1:A6,\"a\",B1:B6,\">10\")", 40)]
    [InlineData("=AVERAGEIFS(B1:B6,A1:A6,\"<>c\")", 16.25)]
    [InlineData("=MAXIFS(B1:B6,A1:A6,\"a\")", 40)]
    [InlineData("=MAXIFS(B1:B6,A1:A6,\"b\")", 20)]
    [InlineData("=MINIFS(B1:B6,A1:A6,\"b\")", -5)]
    [InlineData("=MINIFS(B1:B6,A1:A6,\"a\",B1:B6,\">10\")", 40)]
    [InlineData("=MAXIFS(B1:B6,A1:A6,\"c\")", 0)]
    [InlineData("=MINIFS(B1:B6,A1:A6,\"z\")", 0)]
    [InlineData("=MAXIFS(B:B,A:A,\"b\")", 20)]
    public void Conditional_averages_and_extremes(string formula, double expected)
    {
        var value = Eval(formula);
        Assert.Equal(CellValueKind.Number, value.Kind);
        Assert.Equal(expected, value.AsNumber(), 12);
    }

    [Theory]
    [InlineData("=AVERAGEIF(A1:A6,\"c\",B1:B6)", ErrorKind.Div0)]
    [InlineData("=AVERAGEIF(A1:A6,\"z\",B1:B6)", ErrorKind.Div0)]
    [InlineData("=AVERAGEIFS(B1:B6,A1:A5,\"a\")", ErrorKind.Value)]
    [InlineData("=MAXIFS(B1:B6,A1:A5,\"a\")", ErrorKind.Value)]
    [InlineData("=MINIFS(B1:B6,A1:A6)", ErrorKind.Value)]
    [InlineData("=AVERAGEIF(C1:D1,\"<>e\")", ErrorKind.Div0)]
    [InlineData("=MAXIFS(C1:D1,C1:D1,\"<>e\")", ErrorKind.Div0)]
    public void Conditional_errors(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }

    [Fact]
    public void An_array_of_criteria_gives_an_array()
    {
        var result = Eval("=MAXIFS(B1:B6,A1:A6,{\"a\";\"b\"})").AsArray();
        Assert.Equal(CellValue.Number(40), result[0, 0]);
        Assert.Equal(CellValue.Number(20), result[1, 0]);
    }
}
