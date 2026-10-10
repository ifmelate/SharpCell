using System;
using System.Threading;
using SharpCell.Evaluation;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class WritePlanTests
{
    private static readonly XlsxWriteOptions Default = new();

    private static Workbook Calculated(TestXlsx? xlsx = null)
    {
        var workbook = XlsxReader.Load((xlsx ?? XlsxSourceTests.Template()).Build());
        workbook.Recalculate();
        return workbook;
    }

    [Fact]
    public void A_changed_constant_can_be_saved()
    {
        var workbook = Calculated();
        workbook["S"]["A1"].Value = 5;
        workbook.Recalculate();

        var plan = WritePlan.Create(workbook, Default);

        Assert.Empty(plan.KeptCells);
        Assert.False(plan.FullCalcOnLoad);
    }

    [Fact]
    public void A_workbook_built_in_code_cannot_be_saved()
    {
        var workbook = new Workbook();
        workbook.AddSheet("S");

        Assert.Throws<NotSupportedException>(() => WritePlan.Create(workbook, Default));
    }

    [Theory]
    [MemberData(nameof(XlsxSourceTests.Changes), MemberType = typeof(XlsxSourceTests))]
    public void A_changed_structure_cannot_be_saved(string change)
    {
        var workbook = Calculated();
        XlsxSourceTests.Change(workbook, change);
        workbook.Recalculate();

        Assert.Throws<NotSupportedException>(() => WritePlan.Create(workbook, Default));
    }

    [Theory]
    [InlineData("changed")]
    [InlineData("removed")]
    [InlineData("replaced by a value")]
    [InlineData("added")]
    public void A_formula_change_cannot_be_saved(string change)
    {
        var workbook = Calculated();
        var sheet = workbook["S"];
        switch (change)
        {
            case "changed": sheet["A2"].Formula = "=A1*4"; break;
            case "removed": sheet["A2"].Formula = null; break;
            case "replaced by a value": sheet["A2"].Value = 1; break;
            case "added": sheet["B1"].Formula = "=A1"; break;
        }

        workbook.Recalculate();
        var ex = Assert.Throws<NotSupportedException>(() => WritePlan.Create(workbook, Default));
        Assert.Contains(change == "added" ? "S!B1" : "S!A2", ex.Message);
    }

    [Fact]
    public void A_workbook_not_recalculated_after_loading_cannot_be_saved()
    {
        var workbook = XlsxReader.Load(XlsxSourceTests.Template().Build());

        Assert.Throws<InvalidOperationException>(() => WritePlan.Create(workbook, Default));
    }

    [Fact]
    public void A_cancelled_recalculation_leaves_the_workbook_unsaveable()
    {
        var workbook = Calculated();
        workbook["S"]["A1"].Value = 5;
        Assert.Throws<OperationCanceledException>(() => workbook.Recalculate(new CancellationToken(canceled: true)));

        var ex = Assert.Throws<InvalidOperationException>(() => WritePlan.Create(workbook, Default));
        Assert.IsNotType<XlsxWriteException>(ex);
    }

    internal static TestXlsx WithUnknownFunction() => new TestXlsx().Sheet("S",
        "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"
        + "<row r=\"2\"><c r=\"A2\"><f>NOSUCHFN(A1)</f><v>7</v></c></row>"
        + "<row r=\"3\"><c r=\"A3\"><f>A1*2</f><v>2</v></c></row>");

    [Fact]
    public void Formulas_SharpCell_cannot_calculate_refuse_the_save()
    {
        var workbook = Calculated(WithUnknownFunction());

        var ex = Assert.Throws<XlsxWriteException>(() => WritePlan.Create(workbook, Default));
        var problem = Assert.Single(ex.Problems);
        Assert.Equal("A2", problem.Address);
        Assert.Contains("S!A2", ex.Message);
        Assert.Contains(nameof(XlsxWriteOptions.KeepUncalculated), ex.Message);
    }

    [Fact]
    public void KeepUncalculated_keeps_those_formulas_and_asks_Excel_to_recalculate()
    {
        var workbook = Calculated(WithUnknownFunction());

        var plan = WritePlan.Create(workbook, new XlsxWriteOptions { KeepUncalculated = true });

        Assert.Equal(new CellKey(workbook["S"], 2, 1), Assert.Single(plan.KeptCells));
        Assert.True(plan.FullCalcOnLoad);
    }
}
