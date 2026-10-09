using System;
using System.Linq;
using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Model;

public class SheetLoaderTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public SheetLoaderTests()
    {
        _s = _wb.AddSheet("S");
    }

    private static CellValue N(double value) => CellValue.Number(value);

    private static FormulaNode Parse(string text, int row, int column) => FormulaParser.Parse(text, new CellAddress(row, column));

    private SheetLoader Loader() => new(_s);

    [Fact]
    public void Loaded_formula_shows_the_cached_value_until_recalculation()
    {
        var loader = Loader();
        loader.SetValue(1, 1, N(2));
        loader.SetFormula(1, 2, Parse("=A1*3", 1, 2), LoadedFormulaKind.Legacy, null, N(99));
        loader.Complete();

        Assert.Equal(N(99), _s["B1"].Value);
        Assert.Equal("=A1*3", _s["B1"].Formula);

        _wb.Recalculate();
        Assert.Equal(N(6), _s["B1"].Value);
    }

    [Fact]
    public void Formula_text_is_printed_canonically_without_file_prefixes()
    {
        var loader = Loader();
        loader.SetFormula(1, 1, Parse("=_xlfn.XLOOKUP(1,B1:B2,C1:C2)", 1, 1), LoadedFormulaKind.Legacy, null, N(0));
        loader.Complete();

        Assert.Equal("=XLOOKUP(1,B1:B2,C1:C2)", _s["A1"].Formula);
    }

    [Fact]
    public void One_tree_serves_every_cell_of_a_shared_formula()
    {
        var tree = Parse("=A1*2", 1, 2);
        var loader = Loader();
        loader.SetValue(1, 1, N(1));
        loader.SetValue(2, 1, N(5));
        loader.SetFormula(1, 2, tree, LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.SetFormula(2, 2, tree, LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.Complete();
        _wb.Recalculate();

        Assert.Equal([N(2), N(10)], [_s["B1"].Value, _s["B2"].Value]);
        Assert.Equal("=A2*2", _s["B2"].Formula);
    }

    [Fact]
    public void Legacy_formula_intersects_a_range_result_with_its_row()
    {
        var loader = Loader();
        loader.SetValue(1, 1, N(1));
        loader.SetValue(2, 1, N(2));
        loader.SetValue(3, 1, N(3));
        loader.SetFormula(2, 2, Parse("=A1:A3", 2, 2), LoadedFormulaKind.Legacy, null, N(2));
        loader.Complete();
        _wb.Recalculate();

        Assert.Equal(N(2), _s["B2"].Value);
        Assert.Equal(CellValue.Empty, _s["B3"].Value);
    }

    [Theory]
    [InlineData("=A1:A3*10", 2, 20.0)]
    [InlineData("=-A1:A3", 3, -3.0)]
    [InlineData("=SUM(A1:A3*1)", 2, 2.0)]
    [InlineData("=SUM(A1:A3)", 2, 6.0)]
    [InlineData("=SUM({1,2,3}*2)", 2, 12.0)]
    [InlineData("=ABS(A1:A3)", 3, 3.0)]
    [InlineData("=ROWS(A1:A3)", 2, 3.0)]
    [InlineData("=SUM(A1:A3*B1:B3)", 3, 90.0)]
    public void Legacy_formula_intersects_ranges_where_one_value_is_expected(string formula, int row, double expected)
    {
        var loader = Loader();
        for (var r = 1; r <= 3; r++)
        {
            loader.SetValue(r, 1, N(r));
            loader.SetValue(r, 2, N(r * 10));
        }

        loader.SetFormula(row, 3, Parse(formula, row, 3), LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.Complete();
        _wb.Recalculate();

        Assert.Equal(N(expected), _s[row, 3].Value);
    }

    [Fact]
    public void Legacy_formula_outside_the_rows_of_its_range_is_VALUE()
    {
        var loader = Loader();
        loader.SetValue(1, 1, N(1));
        loader.SetValue(2, 1, N(2));
        loader.SetFormula(5, 3, Parse("=A1:A2*10", 5, 3), LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.SetFormula(5, 4, Parse("=IF(A1:A2>1,1,0)", 5, 4), LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.Complete();
        _wb.Recalculate();

        Assert.Equal(CellValue.Error(ErrorKind.Value), _s["C5"].Value);
        Assert.Equal(CellValue.Error(ErrorKind.Value), _s["D5"].Value);
    }

    [Fact]
    public void Dynamic_formula_keeps_ranges_as_arrays()
    {
        var loader = Loader();
        loader.SetValue(1, 1, N(1));
        loader.SetValue(2, 1, N(2));
        loader.SetFormula(2, 3, Parse("=SUM(A1:A2*10)", 2, 3), LoadedFormulaKind.Dynamic, null, CellValue.Empty);
        loader.Complete();
        _wb.Recalculate();

        Assert.Equal(N(30), _s["C2"].Value);
    }

    [Fact]
    public void Loaded_dynamic_anchor_owns_its_cached_spill()
    {
        var loader = Loader();
        loader.SetValue(1, 1, N(1));
        loader.SetValue(2, 1, N(2));
        loader.SetFormula(1, 2, Parse("=A1:A2*10", 1, 2), LoadedFormulaKind.Dynamic, new Area(1, 2, 2, 2), N(10));
        loader.SetValue(2, 2, N(20));
        loader.Complete();

        Assert.Equal(N(20), _s["B2"].Value);
        Assert.Null(_s["B2"].Formula);

        // The cached spilled value must not block its own anchor.
        _wb.Recalculate();
        Assert.Equal([N(10), N(20)], [_s["B1"].Value, _s["B2"].Value]);

        _s["A2"].Value = 7;
        _wb.Recalculate();
        Assert.Equal(N(70), _s["B2"].Value);
    }

    [Fact]
    public void Loaded_spill_that_shrinks_leaves_no_stale_cells()
    {
        var loader = Loader();
        loader.SetValue(1, 1, N(1));
        loader.SetFormula(1, 2, Parse("=A1:A1", 1, 2), LoadedFormulaKind.Dynamic, new Area(1, 2, 3, 2), N(1));
        loader.SetValue(2, 2, N(8));
        loader.SetValue(3, 2, N(9));
        loader.Complete();
        _wb.Recalculate();

        Assert.Equal(N(1), _s["B1"].Value);
        Assert.Equal([CellValue.Empty, CellValue.Empty], [_s["B2"].Value, _s["B3"].Value]);
    }

    [Fact]
    public void Readers_of_a_loaded_spill_see_recalculated_values()
    {
        var loader = Loader();
        loader.SetValue(1, 1, N(1));
        loader.SetValue(2, 1, N(2));
        loader.SetFormula(1, 2, Parse("=A1:A2+1", 1, 2), LoadedFormulaKind.Dynamic, new Area(1, 2, 2, 2), N(2));
        loader.SetValue(2, 2, N(3));
        loader.SetFormula(1, 3, Parse("=SUM(B1#)", 1, 3), LoadedFormulaKind.Legacy, null, N(5));
        loader.Complete();

        _s["A2"].Value = 10;
        _wb.Recalculate();
        Assert.Equal(N(13), _s["C1"].Value);
    }

    [Fact]
    public void Array_formula_repeats_a_row_over_its_fixed_area()
    {
        LoadArray("={1,2}", new Area(1, 1, 2, 2));
        _wb.Recalculate();

        Assert.Equal([N(1), N(2), N(1), N(2)], [_s["A1"].Value, _s["B1"].Value, _s["A2"].Value, _s["B2"].Value]);
    }

    [Fact]
    public void Array_formula_repeats_a_column_and_a_scalar()
    {
        LoadArray("={1;2}", new Area(1, 1, 2, 2));
        _wb.Recalculate();
        Assert.Equal([N(1), N(1), N(2), N(2)], [_s["A1"].Value, _s["B1"].Value, _s["A2"].Value, _s["B2"].Value]);

        var wb = new Workbook();
        var sheet = wb.AddSheet("S");
        var loader = new SheetLoader(sheet);
        loader.SetFormula(1, 1, Parse("=7", 1, 1), LoadedFormulaKind.Array, new Area(1, 1, 2, 1), N(7));
        loader.Complete();
        wb.Recalculate();
        Assert.Equal([N(7), N(7)], [sheet["A1"].Value, sheet["A2"].Value]);
    }

    [Fact]
    public void Array_formula_pads_a_small_result_with_NA_and_cuts_a_large_one()
    {
        LoadArray("={1,2;3,4}", new Area(1, 1, 3, 3));
        _wb.Recalculate();
        var na = CellValue.Error(ErrorKind.NA);
        Assert.Equal([N(1), N(2), na], [_s["A1"].Value, _s["B1"].Value, _s["C1"].Value]);
        Assert.Equal([na, na, na], [_s["A3"].Value, _s["B3"].Value, _s["C3"].Value]);

        var wb = new Workbook();
        var sheet = wb.AddSheet("S");
        var loader = new SheetLoader(sheet);
        loader.SetFormula(1, 1, Parse("={1,2,3}", 1, 1), LoadedFormulaKind.Array, new Area(1, 1, 1, 2), N(1));
        loader.Complete();
        wb.Recalculate();
        Assert.Equal([N(1), N(2)], [sheet["A1"].Value, sheet["B1"].Value]);
        Assert.Equal(CellValue.Empty, sheet["C1"].Value);
    }

    [Fact]
    public void Array_formula_is_not_blocked_by_its_cached_members_and_never_spills()
    {
        var loader = Loader();
        loader.SetValue(1, 1, N(1));
        loader.SetValue(2, 1, N(2));
        loader.SetFormula(1, 2, Parse("=A1:A2*2", 1, 2), LoadedFormulaKind.Array, new Area(1, 2, 1, 2), N(2));
        loader.Complete();
        _wb.Recalculate();

        // A one-cell array formula keeps the top-left element instead of spilling.
        Assert.Equal(N(2), _s["B1"].Value);
        Assert.Equal(CellValue.Empty, _s["B2"].Value);
    }

    [Fact]
    public void Part_of_an_array_formula_cannot_be_changed()
    {
        LoadArray("={1,2}", new Area(1, 1, 1, 2));
        _wb.Recalculate();

        Assert.Throws<InvalidOperationException>(() => _s["B1"].Value = 5);
        Assert.Throws<InvalidOperationException>(() => _s["B1"].Value = CellValue.Empty);
        Assert.Throws<InvalidOperationException>(() => _s["B1"].Formula = "=1");
        Assert.Equal(N(2), _s["B1"].Value);
    }

    [Fact]
    public void Part_of_an_array_formula_cannot_be_changed_before_the_first_calculation()
    {
        LoadArray("={1,2}", new Area(1, 1, 1, 2), cachedMembers: true);

        Assert.Throws<InvalidOperationException>(() => _s["B1"].Value = 5);
    }

    [Fact]
    public void Replacing_the_anchor_removes_the_whole_array()
    {
        LoadArray("={1,2}", new Area(1, 1, 2, 2));
        _wb.Recalculate();

        _s["A1"].Value = 9;
        _wb.Recalculate();
        Assert.Equal([N(9), CellValue.Empty, CellValue.Empty], [_s["A1"].Value, _s["B1"].Value, _s["B2"].Value]);

        _s["B1"].Value = 3;
        Assert.Equal(N(3), _s["B1"].Value);
    }

    [Fact]
    public void Array_formula_recalculates_and_its_readers_follow()
    {
        var loader = Loader();
        loader.SetValue(1, 1, N(1));
        loader.SetValue(2, 1, N(2));
        loader.SetFormula(1, 2, Parse("=A1:A2*2", 1, 2), LoadedFormulaKind.Array, new Area(1, 2, 2, 2), N(2));
        loader.SetValue(2, 2, N(4));
        loader.SetFormula(1, 3, Parse("=SUM(B1:B2)", 1, 3), LoadedFormulaKind.Legacy, null, N(6));
        loader.Complete();
        _wb.Recalculate();
        Assert.Equal(N(6), _s["C1"].Value);

        _s["A2"].Value = 10;
        _wb.Recalculate();
        Assert.Equal([N(20), N(22)], [_s["B2"].Value, _s["C1"].Value]);
    }

    [Fact]
    public void Spill_reference_to_an_array_formula_covers_its_area()
    {
        // As the corpus shows (dynamic_arrays.xlsx: CONCAT(K1#) over a Ctrl+Shift+Enter K1:K3).
        LoadArray("={1,2,3}", new Area(1, 1, 1, 3));
        _s["D1"].Formula = "=SUM(A1#)";
        _s["E1"].Formula = "=ROWS(A1#)*10+COLUMNS(A1#)";
        _wb.Recalculate();

        Assert.Equal(N(6), _s["D1"].Value);
        Assert.Equal(N(13), _s["E1"].Value);
    }

    [Fact]
    public void Dynamic_spill_is_blocked_by_an_array_formula()
    {
        LoadArray("={1,2}", new Area(2, 1, 2, 2));
        _s["A1"].Formula = "={1;2}";
        _wb.Recalculate();

        Assert.Equal(CellValue.Error(ErrorKind.Spill), _s["A1"].Value);
        Assert.Equal(N(1), _s["A2"].Value);
    }

    [Fact]
    public void Unsupported_formula_keeps_its_text_and_evaluates_to_NAME_with_a_diagnostic()
    {
        var loader = Loader();
        loader.SetUnsupportedFormula(1, 1, "=[1]Sheet1!A1", "links to other workbooks are not supported", N(5));
        loader.SetFormula(1, 2, Parse("=A1+1", 1, 2), LoadedFormulaKind.Legacy, null, N(6));
        loader.Complete();

        Assert.Equal(N(5), _s["A1"].Value);
        Assert.Equal("=[1]Sheet1!A1", _s["A1"].Formula);

        _wb.Recalculate();
        Assert.Equal(CellValue.Error(ErrorKind.Name), _s["A1"].Value);
        Assert.Equal(CellValue.Error(ErrorKind.Name), _s["B1"].Value);
        var diagnostic = Assert.Single(_wb.Diagnostics);
        Assert.Equal(DiagnosticKind.UnsupportedFormula, diagnostic.Kind);
        Assert.Equal("A1", diagnostic.Address);
        Assert.Contains("other workbooks", diagnostic.Message);
    }

    [Fact]
    public void Unsupported_name_evaluates_to_NAME_with_a_diagnostic()
    {
        _wb.DefineUnsupportedName("Ext", "=[1]Sheet1!$A$1", "links to other workbooks are not supported", null);
        _s["A1"].Formula = "=Ext";
        _wb.Recalculate();

        Assert.Equal(CellValue.Error(ErrorKind.Name), _s["A1"].Value);
        Assert.Equal(DiagnosticKind.UnsupportedFormula, Assert.Single(_wb.Diagnostics).Kind);
    }

    [Fact]
    public void Every_loaded_formula_is_calculated_by_the_first_recalculation()
    {
        var loader = Loader();
        for (var row = 1; row <= 100; row++)
            loader.SetFormula(row, 1, Parse("=" + row, row, 1), LoadedFormulaKind.Legacy, null, N(-1));
        loader.Complete();
        _wb.Recalculate();

        Assert.All(Enumerable.Range(1, 100), row => Assert.Equal(N(row), _s[row, 1].Value));
    }

    private void LoadArray(string formula, Area area, bool cachedMembers = false)
    {
        var loader = Loader();
        loader.SetFormula(area.FirstRow, area.FirstColumn, Parse(formula, area.FirstRow, area.FirstColumn), LoadedFormulaKind.Array, area, N(1));
        if (cachedMembers)
        {
            for (var row = area.FirstRow; row <= area.LastRow; row++)
            {
                for (var column = area.FirstColumn; column <= area.LastColumn; column++)
                {
                    if (row != area.FirstRow || column != area.FirstColumn)
                        loader.SetValue(row, column, N(0));
                }
            }
        }

        loader.Complete();
    }
}
