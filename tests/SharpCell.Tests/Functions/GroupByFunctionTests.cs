namespace SharpCell.Tests.Functions;

public class GroupByFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _d;

    public GroupByFunctionTests()
    {
        _d = _wb.AddSheet("Data");
        string[] region = ["North", "South", "North", "South", "North", "South"];
        string[] product = ["A", "B", "A", "C", "B", "A"];
        double[] units = [10, 20, 30, 40, 50, 60];
        _d["A1"].Value = "Region";
        _d["B1"].Value = "Product";
        _d["C1"].Value = "Units";
        for (var i = 0; i < 6; i++)
        {
            _d[i + 2, 1].Value = region[i];
            _d[i + 2, 2].Value = product[i];
            _d[i + 2, 3].Value = units[i];
        }
    }

    [Fact]
    public void PERCENTOF_is_the_sum_of_a_part_over_the_sum_of_all()
    {
        Assert.Equal(60.0 / 210, _wb.Evaluate("=PERCENTOF(Data!C2:C4,Data!C2:C7)").AsNumber(), 12);
        Assert.Equal(0.5, _wb.Evaluate("=PERCENTOF(5,10)").AsNumber(), 12);
    }

    [Theory]
    [InlineData("=PERCENTOF({1,2},{0,0})", "#DIV/0!")]
    [InlineData("=PERCENTOF({1,#N/A},{4})", "#N/A")]
    [InlineData("=PERCENTOF({1},{#NUM!})", "#NUM!")]
    [InlineData("=PERCENTOF(\"x\",4)", "#VALUE!")]
    public void PERCENTOF_errors(string formula, string expected)
    {
        Assert.Equal(expected, _wb.Evaluate(formula).ToString());
    }

    [Fact]
    public void PERCENTOF_ignores_text_in_ranges_and_arrays_as_SUM_does()
    {
        Assert.Equal(0.25, _wb.Evaluate("=PERCENTOF({1,\"x\"},{4})").AsNumber(), 12);
        Assert.Equal(0.0, _wb.Evaluate("=PERCENTOF(Data!B2:B7,Data!C2:C7)").AsNumber(), 12);
    }
}
