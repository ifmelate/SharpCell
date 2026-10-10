using System.Linq;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Evaluation;

public class ChangedCellsTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public ChangedCellsTests() => _s = _wb.AddSheet("S");

    private static string[] Addresses(RecalculationResult result) => [.. result.ChangedCells.Select(c => c.ToString())];

    [Fact]
    public void Formulas_whose_values_changed_are_listed_but_not_the_inputs()
    {
        _s["A1"].Value = 1;
        _s["A2"].Formula = "=A1*2";
        _s["A3"].Formula = "=A2>0";
        Assert.Equal(new[] { "S!A2", "S!A3" }, Addresses(_wb.Recalculate()));

        _s["A1"].Value = 5;
        Assert.Equal(new[] { "S!A2" }, Addresses(_wb.Recalculate()));   // A3 stays TRUE

        _s["A1"].Value = 5;
        Assert.Empty(_wb.Recalculate().ChangedCells);
    }

    [Fact]
    public void A_cell_handle_shows_the_new_value()
    {
        _s["A1"].Formula = "=2+3";
        var cell = Assert.Single(_wb.Recalculate().ChangedCells);
        Assert.Equal(CellValue.Number(5), cell.Value);
    }

    [Fact]
    public void Spilled_cells_that_appear_change_or_disappear_are_listed()
    {
        _s["A1"].Value = 3;
        _s["B1"].Formula = "=SEQUENCE(A1)";
        Assert.Equal(new[] { "S!B1", "S!B2", "S!B3" }, Addresses(_wb.Recalculate()));

        _s["A1"].Value = 2;
        var result = _wb.Recalculate();
        Assert.Equal(new[] { "S!B3" }, Addresses(result));
        Assert.Equal(CellValue.Empty, result.ChangedCells[0].Value);
    }

    [Fact]
    public void A_spill_that_becomes_blocked_lists_its_anchor_and_lost_cells()
    {
        _s["B1"].Formula = "=SEQUENCE(2)";
        _wb.Recalculate();

        _s["B2"].Value = "blocker";
        Assert.Equal(new[] { "S!B1" }, Addresses(_wb.Recalculate()));
        Assert.Equal(CellValue.Error(ErrorKind.Spill), _s["B1"].Value);
    }

    [Fact]
    public void Cells_calculated_by_Evaluate_are_listed_by_the_next_Recalculate()
    {
        _s["A1"].Value = 1;
        _s["A2"].Formula = "=A1+1";
        _wb.Recalculate();
        _s["A1"].Value = 10;
        Assert.Equal(CellValue.Number(11), _wb.Evaluate("=S!A2"));

        Assert.Equal(new[] { "S!A2" }, Addresses(_wb.Recalculate()));
    }

    [Fact]
    public void A_cell_that_changed_and_changed_back_is_not_listed()
    {
        _s["A1"].Value = 1;
        _s["A2"].Formula = "=A1+1";
        _wb.Recalculate();
        _s["A1"].Value = 10;
        _wb.Evaluate("=S!A2");
        _s["A1"].Value = 1;

        Assert.Empty(_wb.Recalculate().ChangedCells);
    }

    [Fact]
    public void Changes_are_kept_when_a_recalculation_is_cancelled()
    {
        _s["A1"].Value = 1;
        _s["A2"].Formula = "=A1+1";
        _wb.Recalculate();
        _s["A1"].Value = 10;
        _wb.Evaluate("=S!A2");

        Assert.Throws<OperationCanceledException>(() => _wb.Recalculate(new CancellationToken(canceled: true)));
        Assert.Equal(new[] { "S!A2" }, Addresses(_wb.Recalculate()));
    }

    [Fact]
    public void Cells_are_in_sheet_order_then_by_row_and_column()
    {
        var first = _wb.AddSheet("First");
        _wb.AddSheet("Middle");
        var last = _wb.AddSheet("Last");
        last["A1"].Formula = "=1";
        first["B2"].Formula = "=1";
        first["A2"].Formula = "=1";
        first["C1"].Formula = "=1";
        _s["Z9"].Formula = "=1";

        Assert.Equal(new[] { "S!Z9", "First!C1", "First!A2", "First!B2", "Last!A1" }, Addresses(_wb.Recalculate()));
    }

    [Fact]
    public void A_new_formula_with_the_same_result_is_listed()
    {
        _s["A1"].Formula = "=1+1";
        _wb.Recalculate();
        _s["A1"].Formula = "=2";

        Assert.Equal(new[] { "S!A1" }, Addresses(_wb.Recalculate()));
    }

    [Fact]
    public void A_constant_typed_over_a_formula_is_not_listed()
    {
        _s["A1"].Formula = "=1+1";
        _s["A2"].Formula = "=A1";
        _wb.Recalculate();
        _s["A1"].Value = 2;

        Assert.Empty(_wb.Recalculate().ChangedCells);
    }

    [Fact]
    public void A_volatile_formula_with_the_same_value_is_not_listed()
    {
        _s["A1"].Formula = "=INDIRECT(\"B1\")";
        _s["B1"].Value = 1;
        _wb.Recalculate();

        Assert.Empty(_wb.Recalculate().ChangedCells);
    }

    [Fact]
    public void The_first_recalculation_after_loading_lists_results_that_differ_from_Excel()
    {
        using var file = new Xlsx.TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>2</v></c><c r=\"B1\"><f>A1*2</f><v>4</v></c><c r=\"C1\"><f>A1*3</f><v>0</v></c></row>").Build();
        var workbook = XlsxReader.Load(file);

        Assert.Equal(new[] { "S!C1" }, Addresses(workbook.Recalculate()));
    }
}
