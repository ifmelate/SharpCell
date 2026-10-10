using System.Linq;

namespace SharpCell.Tests.Evaluation;

public class IterativeCalculationTests
{
    private static (Workbook W, Worksheet S) New(int iterations = 100, double change = 0.001)
    {
        var w = new Workbook { Iteration = new IterationSettings(true, iterations, change) };
        return (w, w.AddSheet("S"));
    }

    [Fact]
    public void Without_iteration_a_cycle_is_zero_and_a_diagnostic_as_before()
    {
        var w = new Workbook();
        var s = w.AddSheet("S");
        s["A1"].Formula = "=A1+1";
        w.Recalculate();
        Assert.Equal(CellValue.Number(0), s["A1"].Value);
        Assert.Equal(DiagnosticKind.CircularReference, Assert.Single(w.Diagnostics).Kind);
    }

    [Fact]
    public void A_self_counter_advances_MaxIterations_per_recalculation()
    {
        var (w, s) = New(iterations: 5, change: 0);
        s["A1"].Formula = "=A1+1";
        Assert.Equal(new[] { "S!A1" }, w.Recalculate().ChangedCells.Select(c => c.ToString()));
        Assert.Equal(CellValue.Number(5), s["A1"].Value);
        w.Recalculate();
        Assert.Equal(CellValue.Number(10), s["A1"].Value);   // a cycle is volatile and continues
        var diagnostic = Assert.Single(w.Diagnostics);
        Assert.Equal(DiagnosticKind.IterationLimitReached, diagnostic.Kind);
        Assert.Equal("A1", diagnostic.Address);
    }

    [Fact]
    public void A_converging_model_settles_within_MaxChange_without_a_diagnostic()
    {
        var (w, s) = New();
        s["A1"].Value = 1000;
        s["A2"].Formula = "=A1+A3";
        s["A3"].Formula = "=(A1+A2)/2*0.05";
        s["B1"].Formula = "=A2";
        w.Recalculate();
        // Closed form: A2 = 1000 * 1.025 / 0.975.
        Assert.Equal(1000 * 1.025 / 0.975, s["A2"].Value.AsNumber(), 2);
        Assert.Equal(s["A2"].Value, s["B1"].Value);
        Assert.Empty(w.Diagnostics);
    }

    [Fact]
    public void A_reader_outside_the_cycle_is_calculated_after_it()
    {
        var (w, s) = New(iterations: 3, change: 0);
        s["A1"].Formula = "=A1+1";
        s["A2"].Formula = "=A1*10";
        w.Recalculate();
        Assert.Equal(CellValue.Number(30), s["A2"].Value);
    }

    [Fact]
    public void A_cycle_through_a_range_iterates()
    {
        var (w, s) = New(iterations: 2, change: 0);
        s["C1"].Formula = "=SUM(C1:C2)+1";
        s["C2"].Value = 5;
        w.Recalculate();
        Assert.Equal(CellValue.Number(12), s["C1"].Value);   // 0 -> 6 -> 12
    }

    [Fact]
    public void Two_cycles_iterate_independently()
    {
        var (w, s) = New(iterations: 2, change: 0);
        s["A1"].Formula = "=A1+1";
        s["C1"].Formula = "=C1+10";
        w.Recalculate();
        Assert.Equal((2.0, 20.0), (s["A1"].Value.AsNumber(), s["C1"].Value.AsNumber()));
    }

    [Fact]
    public void A_cycle_of_two_cells_settles_on_the_same_fixed_point_in_any_order()
    {
        var (w, s) = New();
        s["A1"].Formula = "=B1/2+1";
        s["B1"].Formula = "=A1/2+1";
        w.Recalculate();
        Assert.Equal(2, s["A1"].Value.AsNumber(), 2);
        Assert.Equal(2, s["B1"].Value.AsNumber(), 2);
    }

    [Fact]
    public void A_cycle_reached_through_a_cell_outside_it_includes_every_member()
    {
        // A1 -> B1 -> C1 -> A1 is the cycle; D1 reads it from outside and is calculated first.
        var (w, s) = New(iterations: 4, change: 0);
        s["D1"].Formula = "=A1";
        s["A1"].Formula = "=B1+1";
        s["B1"].Formula = "=C1";
        s["C1"].Formula = "=A1";
        w.Recalculate();
        Assert.Equal(s["A1"].Value, s["D1"].Value);
        Assert.False(w.Calculation.HasDirty);
        // Every member of the cycle iterated; it does not settle, so each has the limit diagnostic.
        Assert.Equal(new[] { "A1", "B1", "C1" }, w.Diagnostics.Where(d => d.Kind == DiagnosticKind.IterationLimitReached).Select(d => d.Address).Order());
        Assert.DoesNotContain(w.Diagnostics, d => d.Kind == DiagnosticKind.CircularReference);
    }

    [Fact]
    public void A_cycle_across_sheets_iterates()
    {
        var (w, s) = New(iterations: 3, change: 0);
        var o = w.AddSheet("Other");
        s["A1"].Formula = "=Other!A1+1";
        o["A1"].Formula = "=S!A1*2";
        w.Recalculate();
        // Each pass keeps Other!A1 = 2 * S!A1 for the S!A1 it read; the exact pair depends on the
        // order of the pass, which the Excel reference workbook pins.
        var a = s["A1"].Value.AsNumber();
        var b = o["A1"].Value.AsNumber();
        Assert.True(b == 2 * a || a == b + 1, $"S!A1 {a}, Other!A1 {b}");
    }

    [Fact]
    public void Cancellation_during_iteration_leaves_the_workbook_usable()
    {
        var (w, s) = New(iterations: 1000, change: 0);
        using var cts = new CancellationTokenSource();
        var calls = 0;
        w.Functions.Add("TICK", _ => { if (++calls == 3) cts.Cancel(); return 1; });
        s["A1"].Formula = "=A1+TICK()";
        Assert.ThrowsAny<OperationCanceledException>(() => w.Recalculate(cts.Token));
        Assert.True(w.Calculation.HasDirty);

        w.Recalculate();
        Assert.False(w.Calculation.HasDirty);
        Assert.Equal(CellValueKind.Number, s["A1"].Value.Kind);
    }

    [Fact]
    public void A_cycle_through_a_spill_is_resolved_without_iterating()
    {
        var (w, s) = New();
        s["A1"].Formula = "=SEQUENCE(2)+B2";
        s["B2"].Formula = "=A2";
        w.Recalculate();
        Assert.Contains(w.Diagnostics, d => d.Kind == DiagnosticKind.CircularReference);
    }

    [Fact]
    public void Precedents_and_dependents_of_a_cycle_member_are_its_last_pass_reads()
    {
        var (w, s) = New(iterations: 2, change: 0);
        s["A1"].Formula = "=B1+1";
        s["B1"].Formula = "=A1";
        w.Recalculate();
        Assert.Equal(new[] { "S!B1" }, s["A1"].Precedents.Select(p => p.ToString()));
        Assert.Equal(new[] { "S!B1" }, s["A1"].Dependents.Select(d => d.ToString()));
    }

    [Fact]
    public void Turning_iteration_off_gives_zero_again()
    {
        var (w, s) = New(iterations: 3, change: 0);
        s["A1"].Formula = "=A1+1";
        w.Recalculate();
        w.Iteration = new IterationSettings();
        w.Recalculate();
        Assert.Equal(CellValue.Number(0), s["A1"].Value);
        Assert.Equal(DiagnosticKind.CircularReference, Assert.Single(w.Diagnostics).Kind);
    }

    [Fact]
    public void A_cell_found_to_read_the_cycle_while_iterating_joins_it()
    {
        // A1 reads B1 and C1, both read A1: the loop first found is C1 -> A1; B1 joins when A1 asks for it.
        var (w, s) = New(iterations: 3, change: 0);
        s["A1"].Formula = "=B1+C1+1";
        s["B1"].Formula = "=A1";
        s["C1"].Formula = "=A1";
        w.Recalculate();

        Assert.False(w.Calculation.HasDirty);
        Assert.DoesNotContain(w.Diagnostics, d => d.Kind == DiagnosticKind.CircularReference);
        Assert.Equal(new[] { "A1", "B1", "C1" }, w.Diagnostics.Select(d => d.Address).Order());
        Assert.Equal(s["B1"].Value, s["C1"].Value);
    }

    [Fact]
    public void A_legacy_array_formula_read_through_its_area_iterates_without_throwing()
    {
        using var file = new Xlsx.TestXlsx { WorkbookTail = "<calcPr iterate=\"1\" iterateCount=\"3\" iterateDelta=\"0\"/>" }.Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><f>A3+1</f></c></row><row r=\"2\"><c r=\"A2\"><f t=\"array\" ref=\"A2:A3\">A1</f></c></row>").Build();
        var w = SharpCell.Xlsx.XlsxReader.Load(file);

        w.Recalculate();
        w.Recalculate();

        var s = w["S"];
        Assert.Equal(CellValueKind.Number, s["A1"].Value.Kind);
        Assert.Equal(s["A2"].Value, s["A3"].Value);                              // the array area repeats its one value
        Assert.Equal(s["A1"].Value, s["A3"].Value);   // A2's area was filled after A1 in the last pass
        Assert.False(w.Calculation.HasDirty);
    }

    [Fact]
    public void A_formula_that_spilled_before_and_now_reads_its_old_spill_iterates_without_throwing()
    {
        var (w, s) = New(iterations: 3, change: 0);
        s["C1"].Value = 1;
        s["A1"].Formula = "=IF(C1=1,SEQUENCE(2),A2+1)";
        w.Recalculate();
        s["C1"].Value = 0;

        w.Recalculate();

        Assert.Equal(CellValueKind.Number, s["A1"].Value.Kind);
        Assert.False(w.Calculation.HasDirty);
    }

    [Fact]
    public void A_total_read_by_every_row_is_found_in_a_few_restarts()
    {
        // A1 is a total every row reads: the loop first found is one row and A1; the other rows join.
        const int rows = 2000;
        var (w, s) = New(iterations: 3, change: 0);
        s["A1"].Formula = $"=SUM(B1:B{rows})/{2 * rows}+1";
        for (var r = 1; r <= rows; r++)
            s[$"B{r}"].Formula = "=A1";

        w.Recalculate();

        Assert.False(w.Calculation.HasDirty);
        Assert.Equal(s["A1"].Value, s[$"B{rows}"].Value);
        Assert.True(w.Calculation.EvaluationCount < 20 * rows, $"{w.Calculation.EvaluationCount} evaluations");
        Assert.True(w.Calculation.IterationRestarts < 5, $"{w.Calculation.IterationRestarts} restarts");
    }
}
