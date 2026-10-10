using System.Linq;

namespace SharpCell.Tests.Evaluation;

public class PrecedentsTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public PrecedentsTests() => _s = _wb.AddSheet("S");

    private static string[] Names(IEnumerable<CellRange> ranges) => [.. ranges.Select(r => r.ToString())];

    private static string[] Names(IEnumerable<Cell> cells) => [.. cells.Select(c => c.ToString())];

    [Fact]
    public void Precedents_are_the_cells_and_ranges_a_formula_read()
    {
        var other = _wb.AddSheet("Other");
        _wb.DefineName("Rate", "=S!$D$1");
        _s["A1"].Formula = "=SUM(B1:B3)+C1+Rate+Other!A1";
        _wb.Recalculate();

        Assert.Equal(new[] { "S!B1:B3", "S!C1", "S!D1", "Other!A1" }, Names(_s["A1"].Precedents));
        Assert.Same(other, _s["A1"].Precedents[3].Worksheet);
    }

    [Fact]
    public void A_branch_that_was_not_taken_is_no_precedent()
    {
        _s["A1"].Value = true;
        _s["B1"].Formula = "=IF(A1, C1, D1)";
        _wb.Recalculate();

        Assert.Equal(new[] { "S!A1", "S!C1" }, Names(_s["B1"].Precedents));
    }

    [Fact]
    public void Precedents_follow_references_built_at_run_time()
    {
        _s["A1"].Formula = "=INDIRECT(\"B5\")+SUM(OFFSET(C1,1,0,2,1))";
        _wb.Recalculate();

        Assert.Equal(new[] { "S!C1", "S!C2:C3", "S!B5" }, Names(_s["A1"].Precedents));
    }

    [Fact]
    public void A_spilled_cell_has_its_anchor_as_precedent()
    {
        _s["A1"].Formula = "=SEQUENCE(3)";
        _wb.Recalculate();

        Assert.Equal(new[] { "S!A1" }, Names(_s["A3"].Precedents));
    }

    [Fact]
    public void Constants_and_empty_cells_have_no_precedents()
    {
        _s["A1"].Value = 1;
        Assert.Empty(_s["A1"].Precedents);
        Assert.Empty(_s["B1"].Precedents);
    }

    [Fact]
    public void A_formula_that_is_out_of_date_has_no_known_precedents()
    {
        _s["A1"].Formula = "=B1";
        Assert.Throws<InvalidOperationException>(() => _s["A1"].Precedents);
    }

    [Fact]
    public void Dependents_are_the_formulas_that_read_a_cell()
    {
        _s["A1"].Formula = "=SUM(B1:B3)";
        _s["E1"].Formula = "=B2*2";
        _s["F1"].Formula = "=B9";
        var other = _wb.AddSheet("Other");
        other["A1"].Formula = "=S!B2";
        _wb.Recalculate();

        Assert.Equal(new[] { "S!A1", "S!E1", "Other!A1" }, Names(_s["B2"].Dependents));
        Assert.Empty(_s["C2"].Dependents);
    }

    [Fact]
    public void An_anchor_has_its_spilled_cells_as_dependents()
    {
        _s["A1"].Formula = "=SEQUENCE(2)";
        _s["B1"].Formula = "=A1";
        _wb.Recalculate();

        Assert.Equal(new[] { "S!B1", "S!A2" }, Names(_s["A1"].Dependents));
    }

    [Fact]
    public void Dependents_are_unknown_while_formulas_are_out_of_date()
    {
        _s["A1"].Formula = "=B1";
        _wb.Recalculate();
        _s["B1"].Value = 2;

        Assert.Throws<InvalidOperationException>(() => _s["B1"].Dependents);
        _wb.Recalculate();
        Assert.Equal(new[] { "S!A1" }, Names(_s["B1"].Dependents));
    }

    [Fact]
    public void Dependents_follow_a_changed_formula()
    {
        _s["A1"].Formula = "=B1";
        _wb.Recalculate();
        _s["A1"].Formula = "=C1";
        _wb.Recalculate();

        Assert.Empty(_s["B1"].Dependents);
        Assert.Equal(new[] { "S!A1" }, Names(_s["C1"].Dependents));
    }
}
