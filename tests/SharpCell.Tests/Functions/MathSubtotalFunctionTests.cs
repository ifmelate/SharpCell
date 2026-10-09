using System;
using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Functions;

public class MathSubtotalFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public MathSubtotalFunctionTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = 1;
        _s["A2"].Value = 2;
        _s["A3"].Value = 3;
        _s["A4"].Value = "x";
        _s["A5"].Value = true;
        _s["A6"].Value = 4;
        _s["A7"].Value = 4;
        _s["B1"].Value = 10;
        _s["B2"].Formula = "=1/0";
        _s["B3"].Value = 30;
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    private void AssertNumber(double expected, string formula)
    {
        var actual = Eval(formula);
        Assert.True(actual.Kind == CellValueKind.Number, $"{formula}: {actual}");
        Assert.Equal(expected, actual.AsNumber(), 1e-12 * Math.Max(1, Math.Abs(expected)));
    }

    [Theory]
    [InlineData(1, 2.8)]
    [InlineData(2, 5)]
    [InlineData(3, 7)]
    [InlineData(4, 4)]
    [InlineData(5, 1)]
    [InlineData(6, 96)]
    [InlineData(7, 1.3038404810405297)]
    [InlineData(8, 1.16619037896906)]
    [InlineData(9, 14)]
    [InlineData(10, 1.7)]
    [InlineData(11, 1.36)]
    public void SUBTOTAL_computes_each_function(int code, double expected)
    {
        AssertNumber(expected, $"=SUBTOTAL({code},A1:A7)");
        AssertNumber(expected, $"=SUBTOTAL({code + 100},A1:A7)");
    }

    [Theory]
    [InlineData("=AGGREGATE(12,0,A1:A7)", 3)]
    [InlineData("=AGGREGATE(13,0,A1:A7)", 4)]
    [InlineData("=AGGREGATE(14,0,A1:A7,2)", 4)]
    [InlineData("=AGGREGATE(15,0,A1:A7,2)", 2)]
    [InlineData("=AGGREGATE(16,0,A1:A7,0.25)", 2)]
    [InlineData("=AGGREGATE(17,0,A1:A7,3)", 4)]
    [InlineData("=AGGREGATE(18,0,A1:A7,0.5)", 3)]
    [InlineData("=AGGREGATE(19,0,A1:A7,1)", 1.5)]
    [InlineData("=AGGREGATE(9,6,B1:B3)", 40)]
    [InlineData("=AGGREGATE(9,7,A1:B3)", 46)]
    [InlineData("=AGGREGATE(14,6,B1:B3/A1:A3,1)", 10)]
    [InlineData("=AGGREGATE(2,4,B1:B3)", 2)]
    public void AGGREGATE_computes_like_Excel(string formula, double expected)
    {
        AssertNumber(expected, formula);
    }

    [Theory]
    [InlineData("=SUBTOTAL(9,B1:B3)", ErrorKind.Div0)]
    [InlineData("=SUBTOTAL(12,A1:A3)", ErrorKind.Value)]
    [InlineData("=SUBTOTAL(0,A1:A3)", ErrorKind.Value)]
    [InlineData("=SUBTOTAL(9,{1,2})", ErrorKind.Value)]
    [InlineData("=SUBTOTAL(1,A4:A5)", ErrorKind.Div0)]
    [InlineData("=SUBTOTAL(7,A1)", ErrorKind.Div0)]
    [InlineData("=AGGREGATE(9,4,B1:B3)", ErrorKind.Div0)]
    [InlineData("=AGGREGATE(20,0,A1:A3)", ErrorKind.Value)]
    [InlineData("=AGGREGATE(9,8,A1:A3)", ErrorKind.Value)]
    [InlineData("=AGGREGATE(14,0,A1:A3)", ErrorKind.Value)]
    [InlineData("=AGGREGATE(14,0,A1:A3,4)", ErrorKind.Num)]
    [InlineData("=AGGREGATE(13,0,A1:A3)", ErrorKind.NA)]
    [InlineData("=AGGREGATE(16,0,A1:A3,1.5)", ErrorKind.Num)]
    [InlineData("=AGGREGATE(18,0,A1:A3,0.1)", ErrorKind.Num)]
    public void Fails_like_Excel(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }

    [Fact]
    public void AGGREGATE_array_form_takes_arrays_in_formulas_from_before_dynamic_arrays()
    {
        var loader = new SheetLoader(_s);
        loader.SetFormula(2, 5, FormulaParser.Parse("=AGGREGATE(15,6,B1:B3/A1:A3,1)", new CellAddress(2, 5)), LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.Complete();
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(10), _s["E2"].Value);
    }

    [Fact]
    public void Nested_totals_are_not_counted_twice()
    {
        _s["C1"].Value = 5;
        _s["C2"].Value = 6;
        _s["C3"].Formula = "=SUBTOTAL(9,C1:C2)";
        _s["C4"].Value = 7;
        _s["C5"].Formula = "=AGGREGATE(9,0,C4)*2";
        _s["C6"].Formula = "=SUBTOTAL(9,C1:C5)";
        _s["C7"].Formula = "=SUM(C1:C5)";
        _s["C8"].Formula = "=AGGREGATE(9,4,C1:C5)";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(18), _s["C6"].Value);
        Assert.Equal(CellValue.Number(43), _s["C7"].Value);
        Assert.Equal(CellValue.Number(43), _s["C8"].Value);
    }
}
