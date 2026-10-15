namespace SharpCell.Tests.Model;

/// <summary>How a sheet prints: orientation, margins, print area, print titles and page breaks.</summary>
public class PrintingTests
{
    [Fact]
    public void A_new_sheet_prints_portrait_with_normal_margins_and_everything()
    {
        var sheet = new Workbook().AddSheet("S");

        Assert.Equal(PageOrientation.Portrait, sheet.PageOrientation);
        Assert.Equal(PageMargins.Normal, sheet.PageMargins);
        Assert.Null(sheet.PrintArea);
        Assert.Null(sheet.PrintTitleRows);
        Assert.Empty(sheet.RowPageBreaks);
    }

    [Fact]
    public void Settings_are_set_and_read_back()
    {
        var sheet = new Workbook().AddSheet("S");

        sheet.PageOrientation = PageOrientation.Landscape;
        sheet.PageMargins = PageMargins.Normal with { Left = 0.5, Footer = 0.2 };
        sheet.PrintArea = "$A$1:$F$40, H1:J5";
        sheet.PrintTitleRows = "$1:$2";
        sheet.PrintTitleColumns = "A:A";
        sheet.AddRowPageBreak(21);
        sheet.AddRowPageBreak(11);
        sheet.AddColumnPageBreak(4);

        Assert.Equal(PageOrientation.Landscape, sheet.PageOrientation);
        Assert.Equal(new PageMargins(0.5, 0.7, 0.75, 0.75, 0.3, 0.2), sheet.PageMargins);
        Assert.Equal("A1:F40,H1:J5", sheet.PrintArea);
        Assert.Equal(("1:2", "A:A"), (sheet.PrintTitleRows, sheet.PrintTitleColumns));
        Assert.Equal([11, 21], sheet.RowPageBreaks);
        Assert.Equal([4], sheet.ColumnPageBreaks);
        sheet.RemoveRowPageBreak(11);
        sheet.PrintArea = null;
        Assert.Equal([21], sheet.RowPageBreaks);
        Assert.Null(sheet.PrintArea);
    }

    [Fact]
    public void Wrong_settings_are_refused()
    {
        var sheet = new Workbook().AddSheet("S");

        Assert.Throws<ArgumentException>(() => sheet.PrintArea = "A1:B2;C3");
        Assert.Throws<ArgumentException>(() => sheet.PrintTitleRows = "A:B");
        Assert.Throws<ArgumentException>(() => sheet.PrintTitleColumns = "1:2");
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddRowPageBreak(1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PageMargins.Normal with { Top = -1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageMargins(double.NaN, 0, 0, 0, 0, 0));
    }

    [Fact]
    public void Settings_follow_inserted_and_deleted_rows_and_columns()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet.PrintArea = "A1:F40";
        sheet.PrintTitleRows = "2:3";
        sheet.PrintTitleColumns = "B:B";
        sheet.AddRowPageBreak(21);
        sheet.AddColumnPageBreak(4);

        sheet.InsertRows(1, 2);
        sheet.DeleteColumns(1);

        Assert.Equal("A3:E42", sheet.PrintArea);
        Assert.Equal(("4:5", "A:A"), (sheet.PrintTitleRows, sheet.PrintTitleColumns));
        Assert.Equal([23], sheet.RowPageBreaks);
        Assert.Equal([3], sheet.ColumnPageBreaks);

        sheet.DeleteRows(23);
        sheet.Range("A1").Insert(InsertShift.Down);
        Assert.Empty(sheet.RowPageBreaks);
        Assert.Equal(("A3:E41", "4:5"), (sheet.PrintArea, sheet.PrintTitleRows));
    }

    [Fact]
    public void A_clone_prints_the_same()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet.PageOrientation = PageOrientation.Landscape;
        sheet.PrintArea = "A1:B2";
        sheet.AddRowPageBreak(5);

        var clone = workbook.Clone()["S"];

        Assert.Equal((PageOrientation.Landscape, "A1:B2", 5), (clone.PageOrientation, clone.PrintArea, clone.RowPageBreaks.Single()));
    }
}
