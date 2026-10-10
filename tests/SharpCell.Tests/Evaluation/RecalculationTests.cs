using System.Globalization;
using SharpCell;
using SharpCell.Functions;

namespace SharpCell.Tests.Evaluation;

public class RecalculationTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public RecalculationTests()
    {
        _s = _wb.AddSheet("Sheet1");
    }

    private static CellValue N(double value) => CellValue.Number(value);

    [Fact]
    public void Spec_example()
    {
        var wb = new Workbook();
        var ws = wb.AddSheet("Sheet1");
        ws["A1"].Value = 2;
        ws["A2"].Formula = "=A1*3";
        wb.Recalculate();
        Assert.Equal(N(6), ws["A2"].Value);
    }

    [Fact]
    public void Values_are_stale_until_recalculation()
    {
        _s["A1"].Value = 2;
        _s["A2"].Formula = "=A1*3";
        Assert.Equal(CellValue.Empty, _s["A2"].Value);
        _wb.Recalculate();
        _s["A1"].Value = 5;
        Assert.Equal(N(6), _s["A2"].Value);
        _wb.Recalculate();
        Assert.Equal(N(15), _s["A2"].Value);
    }

    [Fact]
    public void Only_dependents_of_a_change_are_recalculated()
    {
        _s["A1"].Value = 1;
        _s["B1"].Value = 2;
        _s["A2"].Formula = "=A1+1";
        _s["B2"].Formula = "=B1+1";
        _s["C2"].Formula = "=A2+B2";
        _wb.Recalculate();

        var before = _wb.Calculation.EvaluationCount;
        _s["A1"].Value = 10;
        _wb.Recalculate();
        Assert.Equal(2, _wb.Calculation.EvaluationCount - before);
        Assert.Equal(N(14), _s["C2"].Value);
    }

    [Fact]
    public void Whole_column_range_is_one_interval()
    {
        _s["B1"].Formula = "=SUM(A:A)";
        _wb.Recalculate();
        Assert.Equal(1, _wb.Calculation.Graph.RangeCount);

        _s["A500000"].Value = 7;
        _wb.Recalculate();
        Assert.Equal(N(7), _s["B1"].Value);
    }

    [Fact]
    public void Range_dependents_are_found_through_the_interval_index()
    {
        for (var i = 1; i <= 200; i++)
            _s[i, 3].Formula = $"=SUM(A{i}:B{i + 10})";
        _wb.Recalculate();

        _s["A15"].Value = 1;
        _wb.Recalculate();
        for (var i = 1; i <= 200; i++)
            Assert.Equal(N(i is >= 5 and <= 15 ? 1 : 0), _s[i, 3].Value);
    }

    [Fact]
    public void Chain_of_50000_cells_does_not_use_the_thread_stack()
    {
        const int length = 50_000;
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                // Formulas are entered bottom-up, so calculation starts at the far end of the chain.
                for (var row = length; row >= 2; row--)
                    _s[row, 1].Formula = $"=A{row - 1}+1";
                _s["A1"].Value = 1;
                _wb.Recalculate();
            }
            catch (Exception ex)
            {
                error = ex;
            }
        }, maxStackSize: 512 * 1024);
        thread.Start();
        thread.Join();

        Assert.Null(error);
        Assert.Equal(N(length), _s[length, 1].Value);
    }

    [Fact]
    public void Untaken_branch_is_not_a_dependency()
    {
        _s["A1"].Value = false;
        _s["B1"].Value = 1;
        _s["C1"].Formula = "=IF(A1,B1,0)";
        _wb.Recalculate();

        var before = _wb.Calculation.EvaluationCount;
        _s["B1"].Value = 2;
        _wb.Recalculate();
        Assert.Equal(before, _wb.Calculation.EvaluationCount);

        _s["A1"].Value = true;
        _wb.Recalculate();
        Assert.Equal(N(2), _s["C1"].Value);
        _s["B1"].Value = 3;
        _wb.Recalculate();
        Assert.Equal(N(3), _s["C1"].Value);
    }

    [Fact]
    public void Cells_in_a_cycle_get_zero_and_a_diagnostic()
    {
        _s["A1"].Formula = "=B1+1";
        _s["B1"].Formula = "=A1+1";
        _s["C1"].Formula = "=A1+5";
        _wb.Recalculate();

        Assert.Equal(N(0), _s["A1"].Value);
        Assert.Equal(N(0), _s["B1"].Value);
        Assert.Equal(N(5), _s["C1"].Value);
        var diagnostic = Assert.Single(_wb.Diagnostics);
        Assert.Equal(DiagnosticKind.CircularReference, diagnostic.Kind);
        Assert.Contains("A1", diagnostic.Message);
        Assert.Contains("B1", diagnostic.Message);

        _s["B1"].Value = 10;
        _wb.Recalculate();
        Assert.Equal(N(11), _s["A1"].Value);
        Assert.Equal(N(16), _s["C1"].Value);
        Assert.Empty(_wb.Diagnostics);
    }

    [Fact]
    public void Self_reference_is_a_cycle()
    {
        _s["A1"].Formula = "=A1+1";
        _wb.Recalculate();
        Assert.Equal(N(0), _s["A1"].Value);
        Assert.Equal(DiagnosticKind.CircularReference, Assert.Single(_wb.Diagnostics).Kind);
    }

    [Fact]
    public void Cycle_in_an_untaken_branch_is_not_a_cycle()
    {
        _s["A1"].Formula = "=IF(FALSE,A1,3)";
        _wb.Recalculate();
        Assert.Equal(N(3), _s["A1"].Value);
        Assert.Empty(_wb.Diagnostics);
    }

    [Fact]
    public void Volatile_cells_and_their_dependents_recalculate_every_time()
    {
        _wb.Random = new Random(1);
        _s["A1"].Formula = "=RAND()";
        _s["B1"].Formula = "=A1*10";
        _s["C1"].Formula = "=1+1";
        _wb.Recalculate();
        var first = _s["A1"].Value;

        var before = _wb.Calculation.EvaluationCount;
        _wb.Recalculate();
        Assert.NotEqual(first, _s["A1"].Value);
        Assert.Equal(_s["A1"].Value.AsNumber() * 10, _s["B1"].Value.AsNumber(), 12);
        Assert.Equal(2, _wb.Calculation.EvaluationCount - before);
    }

    [Fact]
    public void Cancellation_leaves_the_workbook_consistent()
    {
        for (var row = 2; row <= 1000; row++)
            _s[row, 1].Formula = $"=A{row - 1}+1";
        _s["A1"].Value = 1;

        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => _wb.Recalculate(cancelled.Token));

        _wb.Recalculate();
        Assert.Equal(N(1000), _s["A1000"].Value);
        Assert.Empty(_wb.Diagnostics);
    }

    [Fact]
    public void Redefining_a_name_recalculates_its_users()
    {
        _wb.DefineName("Rate", "=0.1");
        _s["A1"].Formula = "=100*Rate";
        _s["B1"].Formula = "=Later*2";
        _wb.Recalculate();
        Assert.Equal(N(10), _s["A1"].Value);
        Assert.Equal(CellValue.Error(ErrorKind.Name), _s["B1"].Value);

        _wb.DefineName("Rate", "=0.2");
        _wb.DefineName("Later", "=3");
        _wb.Recalculate();
        Assert.Equal(N(20), _s["A1"].Value);
        Assert.Equal(N(6), _s["B1"].Value);
    }

    [Fact]
    public void Evaluate_calculates_formula_cells_on_demand()
    {
        _s["A1"].Value = 1;
        _s["B1"].Value = 2;
        _s["C1"].Formula = "=A1+B1";
        Assert.Equal(N(30), _wb.Evaluate("=C1*10"));
        Assert.Equal(N(3), _s["C1"].Value);
    }

    [Fact]
    public void Changing_the_culture_recalculates_everything()
    {
        _s["A1"].Formula = "=\"1,5\"+0";
        _wb.Recalculate();
        Assert.Equal(CellValue.Error(ErrorKind.Value), _s["A1"].Value);

        _wb.Culture = CultureInfo.GetCultureInfo("ru-RU");
        _wb.Recalculate();
        Assert.Equal(N(1.5), _s["A1"].Value);
    }

    [Fact]
    public void Adding_a_sheet_resolves_references_to_it()
    {
        _s["A1"].Formula = "=New!A1";
        _wb.Recalculate();
        Assert.Equal(CellValue.Error(ErrorKind.Ref), _s["A1"].Value);

        _wb.AddSheet("New")["A1"].Value = 4;
        _wb.Recalculate();
        Assert.Equal(N(4), _s["A1"].Value);
    }

    [Fact]
    public void Clearing_an_input_updates_dependents()
    {
        _s["A1"].Value = 5;
        _s["B1"].Formula = "=A1";
        _s["C1"].Formula = "=B1*2";
        _wb.Recalculate();
        Assert.Equal(N(10), _s["C1"].Value);

        _s["A1"].Value = CellValue.Empty;
        _wb.Recalculate();
        Assert.Equal(N(0), _s["B1"].Value);

        _s["B1"].Formula = null;
        _s["B1"].Value = 7;
        _wb.Recalculate();
        Assert.Equal(N(14), _s["C1"].Value);
    }

    [Fact]
    public void Cross_sheet_and_3D_dependencies()
    {
        var other = _wb.AddSheet("Other");
        _s["A1"].Formula = "=SUM(Sheet1:Other!B1)";
        _wb.Recalculate();
        other["B1"].Value = 4;
        _wb.Recalculate();
        Assert.Equal(N(4), _s["A1"].Value);
    }

    [Fact]
    public void Function_failure_during_recalculation_points_at_the_cell()
    {
        var registry = new FunctionRegistry();
        registry.Add(new FunctionInfo("BOOM", 0, 0, [ArgumentKind.Value], _ => throw new InvalidOperationException("kaboom")));
        _wb.Registry = registry;
        _s["B2"].Formula = "=BOOM()";
        _wb.Recalculate();

        Assert.Equal(CellValue.Error(ErrorKind.Value), _s["B2"].Value);
        var diagnostic = Assert.Single(_wb.Diagnostics);
        Assert.Same(_s, diagnostic.Sheet);
        Assert.Equal("B2", diagnostic.Address);
    }
}
