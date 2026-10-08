using SharpCell;
using SharpCell.Parsing;

namespace SharpCell.Tests.Evaluation;

public class EvaluatorStackTests
{
    private static CellValue EvaluateOnSmallStack(Workbook wb, string formula)
    {
        CellValue result = default;
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                result = wb.Evaluate(formula);
            }
            catch (Exception ex)
            {
                error = ex;
            }
        }, maxStackSize: 512 * 1024);
        thread.Start();
        thread.Join();
        Assert.Null(error);
        return result;
    }

    [Fact]
    public void Long_operator_chain_evaluates_in_a_loop()
    {
        var wb = new Workbook();
        wb.AddSheet("S")["A1"].Value = 1;
        var formula = "=A1" + string.Concat(Enumerable.Repeat("+A1", 2000));
        Assert.Equal(CellValue.Number(2001), EvaluateOnSmallStack(wb, formula));
    }

    [Theory]
    [InlineData("(", ")")]
    [InlineData("-", "")]
    public void Tree_at_the_depth_limit_evaluates(string open, string close)
    {
        var depth = FormulaLimits.MaxDepth;
        var formula = "=" + string.Concat(Enumerable.Repeat(open, depth)) + "1" + string.Concat(Enumerable.Repeat(close, depth));
        Assert.Equal(CellValue.Number(1), EvaluateOnSmallStack(new Workbook(), formula));
    }

    [Fact]
    public void Percent_chain_at_the_tree_limit_evaluates()
    {
        var formula = "=1" + new string('%', FormulaLimits.MaxTreeDepth - 2);
        Assert.Equal(CellValue.Number(0), EvaluateOnSmallStack(new Workbook(), formula));
    }
}
