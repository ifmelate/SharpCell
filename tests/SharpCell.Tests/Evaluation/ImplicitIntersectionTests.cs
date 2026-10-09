using SharpCell;

namespace SharpCell.Tests.Evaluation;

public class ImplicitIntersectionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public ImplicitIntersectionTests()
    {
        _s = _wb.AddSheet("S");
        var other = _wb.AddSheet("Other");
        for (var row = 1; row <= 3; row++)
        {
            _s[row, 1].Value = row;
            _s[1, row].Value = row;
            other[row, 1].Value = row * 100;
        }
    }

    private CellValue At(string address, string formula)
    {
        _s[address].Formula = formula;
        _wb.Recalculate();
        return _s[address].Value;
    }

    [Theory]
    [InlineData("E2", "=@A1:A3", 2)]
    [InlineData("E3", "=SUM(@A1:A3)", 3)]
    [InlineData("B5", "=@A1:C1", 2)]
    [InlineData("E2", "=@Other!A1:A3", 200)]
    [InlineData("E2", "=@A2", 2)]
    [InlineData("E2", "=@{7,8;9,10}", 7)]
    [InlineData("E2", "=@5", 5)]
    public void Picks_the_cell_in_the_formula_row_or_column(string address, string formula, double expected)
    {
        Assert.Equal(CellValue.Number(expected), At(address, formula));
    }

    [Theory]
    [InlineData("E5", "=@A1:A3")]
    [InlineData("E5", "=@A1:C1")]
    [InlineData("E2", "=@A1:C3")]
    [InlineData("E2", "=@(A1,A2)")]
    public void No_intersection_is_VALUE_error(string address, string formula)
    {
        Assert.Equal(CellValue.Error(ErrorKind.Value), At(address, formula));
    }

    [Fact]
    public void Evaluate_intersects_at_A1()
    {
        Assert.Equal(CellValue.Number(1), _wb.Evaluate("=@A1:A3"));
    }

    [Fact]
    public void Legacy_formula_intersects_its_result_instead_of_spilling()
    {
        _s["E2"].SetFormula("=A1:A3", legacy: true);
        _s["F2"].SetFormula("={7,8}", legacy: true);
        _wb.Recalculate();

        Assert.Equal(CellValue.Number(2), _s["E2"].Value);
        Assert.Equal(CellValue.Empty, _s["E3"].Value);
        Assert.Equal(CellValue.Number(7), _s["F2"].Value);
        Assert.Equal(CellValue.Empty, _s["G2"].Value);
    }
}
