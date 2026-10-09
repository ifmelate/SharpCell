using SharpCell;

namespace SharpCell.Tests.Evaluation;

public class SpillTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public SpillTests()
    {
        _s = _wb.AddSheet("S");
    }

    private static CellValue N(double value) => CellValue.Number(value);

    private static readonly CellValue SpillError = CellValue.Error(ErrorKind.Spill);

    [Fact]
    public void Array_result_spills_into_neighbouring_cells()
    {
        _s["A1"].Formula = "={1,2;3,4}";
        _wb.Recalculate();

        Assert.Equal([N(1), N(2), N(3), N(4)], [_s["A1"].Value, _s["B1"].Value, _s["A2"].Value, _s["B2"].Value]);
        Assert.Null(_s["B1"].Formula);
        Assert.Equal("={1,2;3,4}", _s["A1"].Formula);
    }

    [Fact]
    public void Empty_elements_spill_as_zero()
    {
        _s["D1"].Value = 5;
        _s["A1"].Formula = "=D1:D3";
        _wb.Recalculate();
        Assert.Equal([N(5), N(0), N(0)], [_s["A1"].Value, _s["A2"].Value, _s["A3"].Value]);
    }

    [Fact]
    public void Occupied_cell_blocks_the_spill()
    {
        _s["C1"].Value = "x";
        _s["A1"].Formula = "={1,2,3}";
        _wb.Recalculate();
        Assert.Equal(SpillError, _s["A1"].Value);
        Assert.Equal(CellValue.Empty, _s["B1"].Value);
        Assert.Equal(CellValue.Text("x"), _s["C1"].Value);

        _s["C1"].Value = CellValue.Empty;
        _wb.Recalculate();
        Assert.Equal([N(1), N(2), N(3)], [_s["A1"].Value, _s["B1"].Value, _s["C1"].Value]);
    }

    [Fact]
    public void Formula_cell_blocks_the_spill()
    {
        _s["B1"].Formula = "=1";
        _s["A1"].Formula = "={1,2}";
        _wb.Recalculate();
        Assert.Equal(SpillError, _s["A1"].Value);
        Assert.Equal(N(1), _s["B1"].Value);
    }

    [Fact]
    public void Spill_beyond_the_sheet_edge_is_an_error()
    {
        _s["XFD1"].Formula = "={1,2}";
        _wb.Recalculate();
        Assert.Equal(SpillError, _s["XFD1"].Value);
    }

    [Fact]
    public void Shrinking_spill_clears_cells_it_no_longer_covers()
    {
        _s["D1"].Value = true;
        _s["A1"].Formula = "=IF(D1,{1,2,3},{7})";
        _wb.Recalculate();
        Assert.Equal(N(3), _s["C1"].Value);

        _s["D1"].Value = false;
        _wb.Recalculate();
        Assert.Equal(N(7), _s["A1"].Value);
        Assert.Equal(CellValue.Empty, _s["B1"].Value);
        Assert.Equal(CellValue.Empty, _s["C1"].Value);
        Assert.Equal(2, _s.Store.Count);
    }

    [Fact]
    public void Readers_of_spilled_cells_follow_the_spill()
    {
        _s["E1"].Formula = "=SUM(A1:C1)";
        _s["F1"].Formula = "=B1*10";
        _s["D1"].Value = true;
        _s["A1"].Formula = "=IF(D1,{1,2,3},{7})";
        _wb.Recalculate();
        Assert.Equal(N(6), _s["E1"].Value);
        Assert.Equal(N(20), _s["F1"].Value);

        _s["D1"].Value = false;
        _wb.Recalculate();
        Assert.Equal(N(7), _s["E1"].Value);
        Assert.Equal(N(0), _s["F1"].Value);
    }

    [Fact]
    public void Spill_reference_follows_the_current_spill_area()
    {
        _s["D1"].Value = true;
        _s["A1"].Formula = "=IF(D1,{1;2;3},{7})";
        _s["E1"].Formula = "=SUM(A1#)";
        _s["E2"].Formula = "=ROWS(A1#)";
        _s["E3"].Formula = "=SUM(D1#)";
        _wb.Recalculate();
        Assert.Equal(N(6), _s["E1"].Value);
        Assert.Equal(N(3), _s["E2"].Value);
        Assert.Equal(CellValue.Error(ErrorKind.Ref), _s["E3"].Value);

        _s["D1"].Value = false;
        _wb.Recalculate();
        Assert.Equal(N(7), _s["E1"].Value);
        Assert.Equal(N(1), _s["E2"].Value);
    }

    [Fact]
    public void Typing_into_a_spilled_cell_blocks_the_spill()
    {
        _s["A1"].Formula = "={1,2,3}";
        _wb.Recalculate();
        _s["B1"].Value = 9;
        _wb.Recalculate();

        Assert.Equal(SpillError, _s["A1"].Value);
        Assert.Equal(N(9), _s["B1"].Value);
        Assert.Equal(CellValue.Empty, _s["C1"].Value);
    }

    [Fact]
    public void Removing_the_anchor_formula_removes_the_spill()
    {
        _s["A1"].Formula = "={1,2,3}";
        _s["E1"].Formula = "=SUM(A1:C1)";
        _wb.Recalculate();
        _s["A1"].Formula = null;
        _wb.Recalculate();

        Assert.Equal(CellValue.Empty, _s["B1"].Value);
        Assert.Equal(N(0), _s["E1"].Value);
        Assert.Equal(1, _s.Store.Count);
    }

    [Fact]
    public void Spill_through_a_loop_terminates()
    {
        _s["A1"].Formula = "={1,2}*C1";
        _s["C1"].Formula = "=B1+1";
        _wb.Recalculate();
        Assert.Contains(_wb.Diagnostics, d => d.Kind == DiagnosticKind.CircularReference);
    }

    [Fact]
    public void Evaluate_returns_arrays_without_spilling()
    {
        Assert.Equal(CellValue.Array(new CellValue[,] { { 1, 2 } }), _wb.Evaluate("={1,2}"));
        Assert.Equal(0, _s.Store.Count);
    }
}
