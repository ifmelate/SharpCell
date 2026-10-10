namespace SharpCell.Tests.Evaluation;

public class UnknownFunctionTests
{
    [Fact]
    public void An_unknown_function_is_reported_as_unsupported()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Formula = "=NOSUCHFN(1)";
        workbook.Recalculate();

        Assert.Equal(ErrorKind.Name, sheet["A1"].Value.AsError());
        var diagnostic = Assert.Single(workbook.Diagnostics);
        Assert.Equal(DiagnosticKind.UnsupportedFormula, diagnostic.Kind);
        Assert.Equal("A1", diagnostic.Address);
        Assert.Contains("NOSUCHFN", diagnostic.Message);
    }

    [Fact]
    public void An_undefined_name_is_not_reported()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Formula = "=NOSUCHNAME+1";
        workbook.Recalculate();

        Assert.Equal(ErrorKind.Name, sheet["A1"].Value.AsError());
        Assert.Empty(workbook.Diagnostics);
    }

    [Fact]
    public void An_unknown_function_in_a_branch_not_taken_is_not_reported()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Formula = "=IF(TRUE,1,NOSUCHFN())";
        workbook.Recalculate();

        Assert.Equal(1, sheet["A1"].Value.AsNumber());
        Assert.Empty(workbook.Diagnostics);
    }

    [Fact]
    public void A_name_holding_a_lambda_is_called_without_a_report()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        workbook.DefineName("Twice", "=LAMBDA(x,x*2)");
        sheet["A1"].Formula = "=Twice(4)";
        workbook.Recalculate();

        Assert.Equal(8, sheet["A1"].Value.AsNumber());
        Assert.Empty(workbook.Diagnostics);
    }
}
