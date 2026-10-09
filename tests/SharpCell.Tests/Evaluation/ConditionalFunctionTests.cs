using System.Globalization;
using SharpCell;
using SharpCell.Functions;

namespace SharpCell.Tests.Evaluation;

public class ConditionalFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public ConditionalFunctionTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = "apple";
        _s["A2"].Value = "apricot";
        _s["A3"].Value = "a*b";
        _s["A4"].Value = 10;
        _s["A5"].Value = "10";
        _s["A6"].Value = true;
        _s["B1"].Value = 1;
        _s["B2"].Value = 2;
        _s["B3"].Value = 3;
        _s["B4"].Value = 4;
        _s["B5"].Value = 5;
        _s["B6"].Value = 6;
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    [Theory]
    [InlineData("a*", "apricot", true)]
    [InlineData("a*", "banana", false)]
    [InlineData("A?PLE", "apple", true)]
    [InlineData("a~*b", "a*b", true)]
    [InlineData("a~*b", "axb", false)]
    [InlineData("~~", "~", true)]
    [InlineData("*b", "a*b", true)]
    [InlineData("*", "", true)]
    [InlineData("a*c*e", "abcde", true)]
    [InlineData("a*c*e", "abcdf", false)]
    [InlineData("~a", "~a", true)]
    public void Wildcards(string pattern, string text, bool expected)
    {
        Assert.Equal(expected, Wildcard.Matches(pattern, text, CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData("=COUNTIF(A1:A6,\"a*\")", 3)]
    [InlineData("=COUNTIF(A1:A6,\"a~*b\")", 1)]
    [InlineData("=COUNTIF(A1:A6,10)", 2)]
    [InlineData("=COUNTIF(A1:A6,\"<>10\")", 5)]
    [InlineData("=COUNTIF(A1:A6,\">5\")", 1)]
    [InlineData("=COUNTIF(A1:A6,\"true\")", 1)]
    [InlineData("=COUNTIF(A1:A10,\"\")", 4)]
    [InlineData("=COUNTIF(A1:A10,\"<>\")", 6)]
    [InlineData("=COUNTIF(A1:A10,\"<=\")", 0)]
    [InlineData("=COUNTIF(A1:A10,\">\")", 0)]
    [InlineData("=COUNTIF(A:A,\"<>apple\")", 1048575)]
    [InlineData("=COUNTIFS(A1:A6,\"a*\",B1:B6,\">1\")", 2)]
    [InlineData("=SUMIF(A1:A6,\"a*\",B1)", 6)]
    [InlineData("=SUMIF(B1:B6,\">3\")", 15)]
    [InlineData("=SUMIFS(B1:B6,A1:A6,\"a*\",B1:B6,\"<3\")", 3)]
    public void Counts_and_sums_like_Excel(string formula, double expected)
    {
        Assert.Equal(CellValue.Number(expected), Eval(formula));
    }

    [Fact]
    public void Ranges_of_different_shapes_are_VALUE()
    {
        Assert.Equal(CellValue.Error(ErrorKind.Value), Eval("=COUNTIFS(A1:A6,\"a*\",B1:B5,\">1\")"));
    }

    [Fact]
    public void An_array_of_criteria_gives_an_array()
    {
        _s["D1"].Formula = "=COUNTIF(A1:A6,{\"a*\";10})";
        _wb.Recalculate();
        Assert.Equal([CellValue.Number(3), CellValue.Number(2)], [_s["D1"].Value, _s["D2"].Value]);
    }

    [Fact]
    public void Matching_error_in_the_sum_range_is_the_result()
    {
        _s["B2"].Formula = "=1/0";
        Assert.Equal(CellValue.Error(ErrorKind.Div0), Eval("=SUMIF(A1:A6,\"a*\",B1:B6)"));
        Assert.Equal(CellValue.Number(9), Eval("=SUMIF(A1:A6,10,B1:B6)"));
    }

    [Fact]
    public void Resized_sum_range_is_a_dependency()
    {
        _s["D1"].Formula = "=SUMIF(A1:A6,\"a*\",B1)";
        _wb.Recalculate();
        _s["B3"].Value = 30;
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(33), _s["D1"].Value);
    }
}
