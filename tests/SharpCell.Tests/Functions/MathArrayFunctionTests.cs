using System;
using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Functions;

public class MathArrayFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public MathArrayFunctionTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = 1;
        _s["A2"].Value = 3;
        _s["A3"].Value = 5;
        _s["B1"].Value = 2;
        _s["B2"].Value = "x";
        _s["B3"].Value = 6;
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    private static CellValue Row(params double[] values)
    {
        var array = new CellValue[1, values.Length];
        for (var i = 0; i < values.Length; i++)
            array[0, i] = values[i];
        return CellValue.Array(array);
    }

    private static CellValue Matrix(double[,] values)
    {
        var array = new CellValue[values.GetLength(0), values.GetLength(1)];
        for (var r = 0; r < values.GetLength(0); r++)
        {
            for (var c = 0; c < values.GetLength(1); c++)
                array[r, c] = values[r, c];
        }

        return CellValue.Array(array);
    }

    private CellValue Legacy(string formula, int row = 2, int column = 4)
    {
        var loader = new SheetLoader(_s);
        loader.SetFormula(row, column, FormulaParser.Parse(formula, new CellAddress(row, column)), LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.Complete();
        _wb.Recalculate();
        return _s[row, column].Value;
    }

    [Theory]
    [InlineData("=SUMPRODUCT(A1:A3,B1:B3)", 32)]
    [InlineData("=SUMPRODUCT({1,2,3},{4,-2,-4})", -12)]
    [InlineData("=SUMPRODUCT(A1:A3)", 9)]
    [InlineData("=SUMPRODUCT((A1:A3>1)*A1:A3)", 8)]
    [InlineData("=SUMPRODUCT(A:A,A:A)", 35)]
    [InlineData("=SUMX2MY2(A1:A3,B1:B3)", -14)]
    [InlineData("=SUMX2PY2(A1:A3,B1:B3)", 66)]
    [InlineData("=SUMXMY2(A1:A3,B1:B3)", 2)]
    [InlineData("=SUMX2MY2(A1:A3,{1,2,3})", 21)]
    [InlineData("=MDETERM({1,2,3;3,4,-2;2,3.5,6})", -5.5)]
    [InlineData("=MDETERM({1,2;2,4})", 0)]
    [InlineData("=MDETERM(5)", 5)]
    public void Computes_like_Excel(string formula, double expected)
    {
        var actual = Eval(formula);
        Assert.Equal(CellValueKind.Number, actual.Kind);
        Assert.Equal(expected, actual.AsNumber(), 1e-12 * Math.Max(1, Math.Abs(expected)));
    }

    [Theory]
    [InlineData("=SUMPRODUCT(A1:A3,B1:B2)", ErrorKind.Value)]
    [InlineData("=SUMPRODUCT(A1:A3,{1;2;#DIV/0!})", ErrorKind.Div0)]
    [InlineData("=SUMPRODUCT(,A1:A3)", ErrorKind.Value)]
    [InlineData("=SUMXMY2(A1:A3,B1:B2)", ErrorKind.NA)]
    [InlineData("=MMULT(A1:A3,A1:A3)", ErrorKind.Value)]
    [InlineData("=MMULT(A1:B3,{1;2})", ErrorKind.Value)]
    [InlineData("=MMULT({1,2},{1;#DIV/0!})", ErrorKind.Div0)]
    [InlineData("=MMULT(,)", ErrorKind.Value)]
    [InlineData("=MDETERM({1,2})", ErrorKind.Value)]
    [InlineData("=MDETERM(A:A)", ErrorKind.Value)]
    [InlineData("=MMULT(A:A,A:A)", ErrorKind.Value)]
    [InlineData("=MDETERM({1,\"a\";1,2})", ErrorKind.Value)]
    [InlineData("=MINVERSE({1,2;2,4})", ErrorKind.Num)]
    [InlineData("=MINVERSE({1,2})", ErrorKind.Value)]
    [InlineData("=MUNIT(0)", ErrorKind.Value)]
    [InlineData("=MUNIT(\"a\")", ErrorKind.Value)]
    [InlineData("=SEQUENCE(0)", ErrorKind.Value)]
    [InlineData("=SEQUENCE(-1,2)", ErrorKind.Value)]
    [InlineData("=SEQUENCE(100000,100000)", ErrorKind.Num)]
    [InlineData("=RANDARRAY(2,2,5,1)", ErrorKind.Value)]
    [InlineData("=RANDARRAY(1,1,1.2,1.8,TRUE)", ErrorKind.Value)]
    public void Fails_like_Excel(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }

    [Fact]
    public void Matrix_products_and_inverses()
    {
        Assert.Equal(Matrix(new double[,] { { 19, 22 }, { 43, 50 } }), Eval("=MMULT({1,2;3,4},{5,6;7,8})"));
        Assert.Equal(Matrix(new double[,] { { 9, 12, 15 }, { 19, 26, 33 }, { 29, 40, 51 } }), Eval("=MMULT({1,2;3,4;5,6},{1,2,3;4,5,6})"));
        var inverse = Eval("=MINVERSE({1,2;3,4})").AsArray();
        Assert.Equal(-2, inverse[0, 0].AsNumber(), 1e-15);
        Assert.Equal(1, inverse[0, 1].AsNumber(), 1e-15);
        Assert.Equal(1.5, inverse[1, 0].AsNumber(), 1e-15);
        Assert.Equal(-0.5, inverse[1, 1].AsNumber(), 1e-15);
        Assert.Equal(Matrix(new double[,] { { 1, 0 }, { 0, 1 } }), Eval("=MUNIT(2.9)"));
    }

    [Fact]
    public void MINVERSE_rounds_like_Excel()
    {
        // Excel's own result for this matrix has -4.44E-17 where the exact inverse has 0.
        var inverse = Eval("=MINVERSE({3,6,9;9,12,-6;6,10.5,18})").AsArray();
        Assert.Equal(-4.4408920985006264E-17, inverse[1, 1].AsNumber(), 1e-30);
    }

    [Fact]
    public void Matrix_functions_refuse_cells_that_are_not_numbers()
    {
        Assert.Equal(CellValue.Error(ErrorKind.Value), Eval("=MMULT(A1:B2,{1;1})"));
        Assert.Equal(CellValue.Error(ErrorKind.Value), Eval("=MDETERM(A1:B2)"));
    }

    [Fact]
    public void SEQUENCE_adds_the_step_repeatedly()
    {
        Assert.Equal(Row(1, 2, 3, 4), Eval("=SEQUENCE(1,4)"));
        Assert.Equal(Matrix(new double[,] { { 10 }, { 11 }, { 12 } }), Eval("=SEQUENCE(3,,10)"));
        Assert.Equal(Row(1, 3.5, 6), Eval("=SEQUENCE(,3,,2.5)"));
        Assert.Equal(12.999999999999998, Eval("=SEQUENCE(1,4,12.7,0.1)").AsArray()[0, 3].AsNumber());
    }

    [Fact]
    public void RANDARRAY_uses_the_workbook_generator()
    {
        _wb.Random = new Random(3);
        var expected = new Random(3);
        var values = Eval("=RANDARRAY(2,3,10,20)").AsArray();
        Assert.Equal(2, values.GetLength(0));
        Assert.Equal(3, values.GetLength(1));
        foreach (var value in values)
            Assert.Equal(10 + expected.NextDouble() * 10, value.AsNumber(), 1e-12);

        foreach (var value in Eval("=RANDARRAY(5,5,1,3,TRUE)").AsArray())
            Assert.Contains(value.AsNumber(), new[] { 1.0, 2.0, 3.0 });
        Assert.Equal(CellValueKind.Number, Eval("=RANDARRAY()").AsArray()[0, 0].Kind);
    }

    [Fact]
    public void SUMPRODUCT_takes_arrays_in_formulas_from_before_dynamic_arrays()
    {
        // Old Excel calculated array parameters as arrays without Ctrl+Shift+Enter.
        Assert.Equal(CellValue.Number(8), Legacy("=SUMPRODUCT((A1:A3>1)*A1:A3)"));
        Assert.Equal(CellValue.Number(-14), Legacy("=SUMX2MY2(A1:A3*1,B1:B3)"));
        Assert.Equal(CellValue.Number(0), Legacy("=MDETERM(A1:B1*{1;2})"));
    }

    [Fact]
    public void SUMPRODUCT_sees_errors_where_another_range_is_empty()
    {
        _s["C5"].Formula = "=1/0";
        Assert.Equal(CellValue.Error(ErrorKind.Div0), Eval("=SUMPRODUCT(A1:A5,C1:C5)"));
    }
}
