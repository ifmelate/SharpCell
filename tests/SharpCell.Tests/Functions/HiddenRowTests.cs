using System;
using SharpCell.Functions;

namespace SharpCell.Tests.Functions;

public class HiddenRowTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public HiddenRowTests()
    {
        _s = _wb.AddSheet("S");
        double[] values = [1, 2, 4, 8, 16];
        for (var i = 0; i < values.Length; i++)
            _s[i + 1, 1].Value = values[i];
        _s.SetRowHidden(3, true);
    }

    private double Calc(string formula) => _wb.Evaluate(formula).AsNumber();

    [Theory]
    [InlineData("=SUBTOTAL(9,A1:A5)", 31)]
    [InlineData("=SUBTOTAL(109,A1:A5)", 27)]
    [InlineData("=AGGREGATE(9,0,A1:A5)", 31)]
    [InlineData("=AGGREGATE(9,1,A1:A5)", 27)]
    [InlineData("=AGGREGATE(9,4,A1:A5)", 31)]
    [InlineData("=AGGREGATE(9,5,A1:A5)", 27)]
    [InlineData("=AGGREGATE(9,6,A1:A5)", 31)]
    [InlineData("=AGGREGATE(9,7,A1:A5)", 27)]
    public void Without_a_filter_hidden_rows_are_hidden_by_hand(string formula, double expected)
    {
        Assert.Equal(expected, Calc(formula));
    }

    [Theory]
    [InlineData("=SUBTOTAL(9,A1:A5)", 27)]
    [InlineData("=SUBTOTAL(109,A1:A5)", 27)]
    [InlineData("=AGGREGATE(9,4,A1:A5)", 31)]
    [InlineData("=AGGREGATE(9,5,A1:A5)", 27)]
    public void With_a_filter_every_hidden_row_is_filtered_out(string formula, double expected)
    {
        _s.FilterMode = true;
        Assert.Equal(expected, Calc(formula));
    }

    [Fact]
    public void The_referenced_sheet_decides()
    {
        var other = _wb.AddSheet("Other");
        other["A1"].Formula = "=SUBTOTAL(9,S!A1:A5)";
        other["A2"].Formula = "=SUBTOTAL(109,S!A1:A5)";
        other.FilterMode = true;
        other.SetRowHidden(1, true);
        other.SetRowHidden(2, true);
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(31), other["A1"].Value);
        Assert.Equal(CellValue.Number(27), other["A2"].Value);
    }

    [Fact]
    public void Showing_a_row_takes_effect_at_the_next_recalculation()
    {
        _s["C1"].Formula = "=SUBTOTAL(109,A1:A5)";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(27), _s["C1"].Value);

        _s.SetRowHidden(3, false);
        Assert.False(_s.IsRowHidden(3));
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(31), _s["C1"].Value);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1_048_577)]
    public void Row_must_be_on_the_sheet(int row)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _s.IsRowHidden(row));
        Assert.Throws<ArgumentOutOfRangeException>(() => _s.SetRowHidden(row, true));
    }

    [Fact]
    public void Subtotal_and_aggregate_are_no_longer_deviations()
    {
        Assert.True(_wb.Registry.TryGet("SUBTOTAL", out var subtotal));
        Assert.True(_wb.Registry.TryGet("AGGREGATE", out var aggregate));
        Assert.Equal(FunctionStatus.Implemented, subtotal!.Status);
        Assert.Equal(FunctionStatus.Implemented, aggregate!.Status);
    }
}
