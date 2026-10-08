using SharpCell;
using SharpCell.Parsing;

namespace SharpCell.Tests.Evaluation;

/// <summary>Regressions from the stage 2 review.</summary>
public class ReviewRegressionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public ReviewRegressionTests()
    {
        _s = _wb.AddSheet("S");
    }

    private static CellValue N(double value) => CellValue.Number(value);

    [Theory]
    [InlineData("=IFERROR(B1,C1)")]
    [InlineData("=IF(ISERROR(B1),C1,B1)")]
    [InlineData("=CHOOSE(IF(ISERROR(B1),2,1),B1,C1)")]
    public void Dirty_selector_does_not_create_a_false_cycle(string formula)
    {
        _s["A1"].Formula = formula;
        _s["B1"].Formula = "=1";
        _s["C1"].Formula = "=A1";
        _wb.Recalculate();

        Assert.Equal(N(1), _s["A1"].Value);
        Assert.Equal(N(1), _s["C1"].Value);
        Assert.Empty(_wb.Diagnostics);
    }

    [Fact]
    public void Dirty_selector_does_not_create_a_false_cycle_in_Evaluate()
    {
        _s["B1"].Formula = "=1";
        _s["C1"].Formula = "=IFERROR(B1,A1)";
        _s["A1"].Formula = "=C1";
        Assert.Equal(N(1), _wb.Evaluate("=S!A1"));
        Assert.Empty(_wb.Diagnostics);
    }

    [Fact]
    public void Total_above_dirty_formulas_collects_all_inputs_at_once()
    {
        const int n = 2000;
        _s["C1"].Formula = $"=SUM(B2:B{n + 1})";
        for (var row = 2; row <= n + 1; row++)
        {
            _s[row, 1].Value = row;
            _s[row, 2].Formula = $"=A{row}*2";
        }

        _wb.Recalculate();
        Assert.Equal(N(2.0 * Enumerable.Range(2, n).Sum()), _s["C1"].Value);
        Assert.True(_wb.Calculation.EvaluationCount <= n + 10, $"{_wb.Calculation.EvaluationCount} evaluations");
    }

    [Fact]
    public void AND_over_dirty_formulas_collects_all_inputs_at_once()
    {
        const int n = 500;
        _s["C1"].Formula = $"=AND(B1:B{n})";
        for (var row = 1; row <= n; row++)
            _s[row, 2].Formula = "=TRUE";

        _wb.Recalculate();
        Assert.Equal(CellValue.True, _s["C1"].Value);
        Assert.True(_wb.Calculation.EvaluationCount <= n + 10, $"{_wb.Calculation.EvaluationCount} evaluations");
    }

    [Fact]
    public void Huge_broadcast_is_an_error_value_and_other_cells_still_compute()
    {
        _s["A1"].Formula = "=B1:B1048576+C1:XFD1";
        _s["A2"].Formula = "=1+1";
        _wb.Recalculate();

        Assert.Equal(CellValue.Error(ErrorKind.Num), _s["A1"].Value);
        Assert.Equal(N(2), _s["A2"].Value);
    }

    [Fact]
    public void Text_longer_than_Excel_allows_is_VALUE_error()
    {
        _s["A1"].Value = new string('x', 20_000);
        _s["A2"].Formula = "=A1&A1";
        _s["A3"].Formula = "=A1&\"y\"";
        _wb.Recalculate();

        Assert.Equal(CellValue.Error(ErrorKind.Value), _s["A2"].Value);
        Assert.Equal(20_001, _s["A3"].Value.AsText().Length);
    }

    [Fact]
    public void Persistent_cycle_stays_in_diagnostics()
    {
        _s["A1"].Formula = "=B1+RAND()";
        _s["B1"].Formula = "=A1";
        _wb.Recalculate();
        Assert.Single(_wb.Diagnostics);

        _wb.Recalculate();
        Assert.Equal(DiagnosticKind.CircularReference, Assert.Single(_wb.Diagnostics).Kind);

        _s["B1"].Value = 1;
        _wb.Recalculate();
        Assert.Empty(_wb.Diagnostics);
    }

    [Fact]
    public void Diagnostics_from_Evaluate_belong_to_no_cell_and_do_not_accumulate()
    {
        var registry = new SharpCell.Functions.FunctionRegistry();
        registry.Add(new SharpCell.Functions.FunctionInfo("BOOM", 0, 0, [SharpCell.Functions.ArgumentKind.Value],
            _ => throw new InvalidOperationException("kaboom")));
        _wb.Functions = registry;
        _s["A1"].Formula = "=A1";
        _wb.Recalculate();
        Assert.Single(_wb.Diagnostics);

        _wb.Evaluate("=BOOM()");
        _wb.Evaluate("=BOOM()");
        Assert.Equal(2, _wb.Diagnostics.Count);
        var detached = Assert.Single(_wb.Diagnostics, d => d.Kind == DiagnosticKind.FunctionFailure);
        Assert.Null(detached.Sheet);
        Assert.Null(detached.Address);
    }

    [Fact]
    public void Deeply_nested_names_do_not_overflow_the_stack()
    {
        const int levels = 64;
        var nesting = 20;
        _wb.DefineName("lvl_0", "=1");
        for (var k = 1; k < levels; k++)
            _wb.DefineName($"lvl_{k}", "=" + string.Concat(Enumerable.Repeat("SUM(", nesting)) + $"lvl_{k - 1}" + new string(')', nesting));

        CellValue result = default;
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                result = _wb.Evaluate($"=lvl_{levels - 1}");
            }
            catch (Exception ex)
            {
                error = ex;
            }
        }, maxStackSize: 1024 * 1024);
        thread.Start();
        thread.Join();

        Assert.Null(error);
        Assert.True(result == N(1) || result == CellValue.Error(ErrorKind.Num), result.ToString());
    }
}
