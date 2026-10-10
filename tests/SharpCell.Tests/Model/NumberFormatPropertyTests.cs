namespace SharpCell.Tests.Model;

public class NumberFormatPropertyTests
{
    private static Worksheet Sheet() => new Workbook().AddSheet("S");

    [Fact]
    public void A_cell_is_General_by_default_and_keeps_a_code_as_written()
    {
        var sheet = Sheet();
        Assert.Equal("General", sheet["A1"].NumberFormat);
        sheet["A1"].NumberFormat = "#,##0.00";
        Assert.Equal("#,##0.00", sheet["A1"].NumberFormat);
        sheet["A1"].NumberFormat = "general";
        Assert.Equal("General", sheet["A1"].NumberFormat);
        sheet["A1"].NumberFormat = "0%";
        sheet["A1"].NumberFormat = "";
        Assert.Equal("General", sheet["A1"].NumberFormat);
    }

    [Fact]
    public void A_format_survives_value_and_formula_changes_and_an_empty_cell_may_carry_one()
    {
        var sheet = Sheet();
        sheet["A1"].NumberFormat = "0.00";
        sheet["A1"].Value = 1;
        sheet["A1"].Formula = "=A1";
        sheet["A1"].Value = CellValue.Empty;
        Assert.Equal("0.00", sheet["A1"].NumberFormat);
    }

    [Fact]
    public void A_format_alone_is_no_content()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A2"].NumberFormat = "0.00";
        sheet["A1"].Formula = "=SEQUENCE(2)";
        sheet["B1"].Formula = "=COUNTA(A2:A9)+COUNTBLANK(A3:A4)";
        workbook.Recalculate();

        Assert.Equal(CellValue.Number(2), sheet["A2"].Value);   // a format does not block a spill
        Assert.Equal(CellValue.Number(3), sheet["B1"].Value);
        sheet["A1"].Formula = null;
        workbook.Recalculate();
        Assert.Equal("0.00", sheet["A2"].NumberFormat);           // nor does a spill leaving clear it
        Assert.DoesNotContain(sheet.Cells, c => c.Address == "A2");
    }

    [Fact]
    public void An_invalid_code_is_rejected()
    {
        Assert.Throws<ArgumentException>(() => Sheet()["A1"].NumberFormat = "0;0;0;0;0");
        Assert.Throws<ArgumentNullException>(() => Sheet()["A1"].NumberFormat = null!);
    }

    [Fact]
    public void Setting_a_format_does_not_make_formulas_out_of_date()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Formula = "=1";
        workbook.Recalculate();
        sheet["A1"].NumberFormat = "0.00";
        Assert.False(workbook.Calculation.HasDirty);
    }

    [Fact]
    public void A_clone_keeps_formats_of_its_own()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].NumberFormat = "0%";
        var clone = workbook.Clone();
        clone["S"]["A1"].NumberFormat = "0.0";

        Assert.Equal("0.0", clone["S"]["A1"].NumberFormat);
        Assert.Equal("0%", sheet["A1"].NumberFormat);
    }
}
