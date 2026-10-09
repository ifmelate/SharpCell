using SharpCell.Sample;
using SharpCell.Tests.Xlsx;

namespace SharpCell.Tests.Docs;

public class SampleTests
{
    public static string BudgetPath { get; } =
        Path.GetFullPath(Path.Combine(CorpusFiles.Root, "..", "..", "samples", "SharpCell.Sample", "budget.xlsx"));

    [Fact]
    public void The_sample_prints_the_budget_before_and_after_the_change()
    {
        var output = new StringWriter();
        BudgetDemo.Run(BudgetPath, output);

        Assert.Equal("""
            Total per month: 1810
            Left after expenses: 1190
            After rent change to 1350:
            Total per month: 1960
            Left after expenses: 1040
            Food per month: 450
            Average expense: 490
            Sequence: {1;2;3}

            """, output.ToString().ReplaceLineEndings("\n"));
    }
}
