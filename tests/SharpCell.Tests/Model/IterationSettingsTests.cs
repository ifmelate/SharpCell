namespace SharpCell.Tests.Model;

public class IterationSettingsTests
{
    [Fact]
    public void Defaults_match_Excel()
    {
        var settings = new Workbook().Iteration;
        Assert.False(settings.Enabled);
        Assert.Equal(100, settings.MaxIterations);
        Assert.Equal(0.001, settings.MaxChange);
        Assert.Equal(new IterationSettings(), settings);
    }

    [Theory]
    [InlineData(0, 0.001)]
    [InlineData(32768, 0.001)]
    [InlineData(10, -1)]
    [InlineData(10, double.NaN)]
    [InlineData(10, double.PositiveInfinity)]
    public void Limits_are_Excel_s(int iterations, double change)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new IterationSettings(true, iterations, change));
    }

    [Fact]
    public void Changing_the_settings_makes_every_formula_out_of_date()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Formula = "=1";
        workbook.Recalculate();
        workbook.Iteration = new IterationSettings(Enabled: true);
        Assert.True(workbook.Calculation.HasDirty);
        Assert.Throws<ArgumentNullException>(() => workbook.Iteration = null!);
    }

    [Fact]
    public void A_clone_keeps_the_settings()
    {
        var workbook = new Workbook { Iteration = new IterationSettings(true, 5, 0) };
        Assert.Equal(workbook.Iteration, workbook.Clone().Iteration);
    }

    [Fact]
    public void The_settings_cannot_change_from_a_custom_function()
    {
        var workbook = new Workbook();
        workbook.Functions.Add("SET", _ => { workbook.Iteration = new IterationSettings(Enabled: true); return 0; });
        Assert.Equal(CellValue.Error(ErrorKind.Value), workbook.Evaluate("=SET()"));
        Assert.False(workbook.Iteration.Enabled);
    }
}
