using System.Linq;

namespace SharpCell.Tests.Model;

public class SaveSupportTests
{
    [Fact]
    public void Dirty_formulas_are_known_until_recalculation()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        Assert.False(workbook.Calculation.HasDirty);

        sheet["A1"].Formula = "=1+1";
        Assert.True(workbook.Calculation.HasDirty);

        workbook.Recalculate();
        Assert.False(workbook.Calculation.HasDirty);

        sheet["B1"].Value = 3;
        Assert.False(workbook.Calculation.HasDirty);   // no formula reads B1
        sheet["C1"].Formula = "=B1";
        workbook.Recalculate();
        sheet["B1"].Value = 4;
        Assert.True(workbook.Calculation.HasDirty);
    }

    [Fact]
    public void A_cell_calculated_by_Evaluate_is_no_longer_dirty()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Formula = "=2";
        workbook.Evaluate("=S!A1");

        Assert.False(workbook.Calculation.HasDirty);
    }

    [Fact]
    public void Hidden_rows_and_names_can_be_listed()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet.SetRowHidden(3, true);
        sheet.SetRowHidden(1, true);
        workbook.DefineName("Rate", "=0.1");
        workbook.DefineName("Local", "=1", sheet);

        Assert.Equal(new[] { 1, 3 }, sheet.HiddenRows.Order());
        Assert.Equal(new[] { "LOCAL", "RATE" }, workbook.Names.All.Select(n => n.Name).Order());
        Assert.Null(workbook.Source);
    }
}
