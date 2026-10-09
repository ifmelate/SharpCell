using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Functions;
using SharpCell.Parsing;

namespace SharpCell.Tests.Evaluation;

/// <summary>Where formulas without the dynamic array flag reduce ranges by implicit intersection.</summary>
public class LegacyArgumentTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public LegacyArgumentTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = 0;
        _s["A2"].Value = 1;
        _s["A3"].Value = 0;
        _s["B1"].Value = 10;
        _s["B2"].Value = 20;
        _s["B3"].Value = 30;
    }

    private static CellValue N(double value) => CellValue.Number(value);

    private CellValue Legacy(string formula, int row = 2, int column = 3)
    {
        var loader = new SheetLoader(_s);
        loader.SetFormula(row, column, FormulaParser.Parse(formula, new CellAddress(row, column)), LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.Complete();
        _wb.Recalculate();
        return _s[row, column].Value;
    }

    [Fact]
    public void IF_condition_is_intersected()
    {
        Assert.Equal(CellValue.Text("y"), Legacy("=IF(A1:A3,\"y\",\"n\")"));
    }

    [Fact]
    public void IF_condition_outside_the_range_rows_is_VALUE()
    {
        Assert.Equal(CellValue.Error(ErrorKind.Value), Legacy("=IF(A1:A3,1,2)", row: 5));
    }

    [Fact]
    public void IFERROR_value_is_intersected()
    {
        Assert.Equal(N(1), Legacy("=IFERROR(A1:A3,99)"));
    }

    [Fact]
    public void IF_branch_range_is_intersected_by_the_result()
    {
        Assert.Equal(N(20), Legacy("=IF(TRUE,B1:B3)"));
    }

    [Fact]
    public void Array_context_arguments_keep_ranges_whole_in_legacy_formulas()
    {
        var registry = new FunctionRegistry();
        MathFunctions.Register(registry);
        LogicalFunctions.Register(registry);
        registry.Add(new FunctionInfo("ARRAYSUM", 1, 1, [ArgumentKind.ArrayContext], call =>
        {
            var total = 0.0;
            var value = Evaluator.ToValue(call[0], call.Context);
            foreach (var element in value.Kind == CellValueKind.Array ? value.AsArray() : new[,] { { value } })
                total += element.Kind == CellValueKind.Number ? element.AsNumber() : 0;
            return CellValue.Number(total);
        }));
        _wb.Functions = registry;

        Assert.Equal(N(20), Legacy("=ARRAYSUM((A1:A3>0)*B1:B3)"));
        Assert.Equal(N(60), Legacy("=ARRAYSUM(B1:B3*1)", row: 3, column: 4));

        // IF's condition is intersected even inside an array parameter, as SUMPRODUCT(IF(...)) was.
        Assert.Equal(N(0), Legacy("=ARRAYSUM(IF(A1:A3>0,B1:B3,0))", row: 1, column: 7));
        Assert.Equal(N(60), Legacy("=ARRAYSUM(IF(A1:A3>0,B1:B3,0))", row: 2, column: 7));

        // Outside the array context the same formula is intersected again.
        Assert.Equal(N(20), Legacy("=SUM(B1:B3*1)", row: 2, column: 5));
        Assert.Equal(N(80), Legacy("=ARRAYSUM(B1:B3*1)+B1:B3", row: 2, column: 6));
    }
}
