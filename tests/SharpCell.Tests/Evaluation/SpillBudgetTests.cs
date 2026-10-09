using System.Diagnostics;
using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Evaluation;

/// <summary>All spills and array formulas of a workbook share one cell budget.</summary>
public class SpillBudgetTests
{
    private readonly Workbook _wb = new() { MaxSpillCells = 10 };
    private readonly Worksheet _s;

    public SpillBudgetTests()
    {
        _s = _wb.AddSheet("S");
    }

    private static CellValue N(double value) => CellValue.Number(value);

    [Fact]
    public void Spill_over_the_budget_is_SPILL_with_a_diagnostic()
    {
        _s["A1"].Formula = "=Z1:Z20";
        _wb.Recalculate();

        Assert.Equal(CellValue.Error(ErrorKind.Spill), _s["A1"].Value);
        Assert.Equal(CellValue.Empty, _s["A2"].Value);
        var diagnostic = Assert.Single(_wb.Diagnostics);
        Assert.Equal(DiagnosticKind.LimitExceeded, diagnostic.Kind);
    }

    [Fact]
    public void Spills_share_the_budget_and_release_it()
    {
        _s["A1"].Formula = "={1;2;3;4;5;6}";
        _s["C1"].Formula = "={1;2;3;4;5;6}";
        _wb.Recalculate();
        var spilled = new[] { _s["A1"].Value, _s["C1"].Value };
        Assert.Single(spilled, v => v == CellValue.Error(ErrorKind.Spill));

        _s["A1"].Value = 1;
        _s["C1"].Formula = "={1;2;3;4;5;6;7;8;9}";
        _wb.Recalculate();
        Assert.Equal(N(9), _s["C9"].Value);
    }

    [Fact]
    public void Array_formula_is_protected_by_its_area_without_stored_members()
    {
        var loader = new SheetLoader(_s);
        Assert.True(loader.SetFormula(1, 1, FormulaParser.Parse("={1;2;3}", new CellAddress(1, 1)), LoadedFormulaKind.Array, new Area(1, 1, 3, 1), N(1)));
        loader.Complete();

        Assert.Equal(1, _s.Store.Count);
        Assert.Throws<InvalidOperationException>(() => _s["A3"].Value = 5);

        _wb.Recalculate();
        Assert.Equal(N(3), _s["A3"].Value);

        _s["A1"].Value = 0;
        _s["A3"].Value = 5;
        Assert.Equal(N(5), _s["A3"].Value);
    }

    [Fact]
    public void Loading_refuses_areas_over_the_budget()
    {
        var loader = new SheetLoader(_s);
        Assert.False(loader.SetFormula(1, 1, FormulaParser.Parse("=1", new CellAddress(1, 1)), LoadedFormulaKind.Array, new Area(1, 1, 20, 1), N(1)));
        loader.Complete();

        _s["A5"].Value = 3;
        Assert.Equal(N(3), _s["A5"].Value);
    }

    [Fact]
    public void Removing_a_wide_array_formula_is_fast()
    {
        var wb = new Workbook();
        var sheet = wb.AddSheet("S");
        var loader = new SheetLoader(sheet);
        loader.SetFormula(1, 1, FormulaParser.Parse("=1", new CellAddress(1, 1)), LoadedFormulaKind.Array,
            new Area(1, 1, 20, CellAddress.MaxColumn), N(1));
        loader.Complete();
        wb.Recalculate();

        var watch = Stopwatch.StartNew();
        sheet["A1"].Value = 0;
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(5), $"Took {watch.Elapsed.TotalSeconds:0.0} s.");
        Assert.Equal(1, sheet.Store.Count);
    }
}
