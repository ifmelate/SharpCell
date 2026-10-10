using System.Globalization;

namespace SharpCell.Tests.Model;

public class CellTextTests
{
    private static (Workbook Workbook, Worksheet Sheet) New()
    {
        var workbook = new Workbook();
        return (workbook, workbook.AddSheet("S"));
    }

    [Theory]
    [InlineData(1234.5678, "0.00", "1234.57")]
    [InlineData(1234.5678, "#,##0.00", "1,234.57")]
    [InlineData(0.256, "0%", "26%")]
    [InlineData(46096.5, "yyyy-mm-dd", "2026-03-15")]
    [InlineData(46096.5, "h:mm AM/PM", "12:00 PM")]
    [InlineData(-5, "0.00;(0.00)", "(5.00)")]
    public void A_number_is_formatted_by_the_cell_format(double value, string code, string expected)
    {
        var (_, sheet) = New();
        sheet["A1"].Value = value;
        sheet["A1"].NumberFormat = code;
        Assert.Equal(expected, sheet["A1"].Text);
    }

    [Theory]
    [InlineData(1234.5678, "1234.5678")]
    [InlineData(1.5e21, "1.5E+21")]
    [InlineData(123456789012, "1.23457E+11")]
    [InlineData(0.1 + 0.2, "0.3")]
    public void General_text_matches_Excel_general_format(double value, string expected)
    {
        var (_, sheet) = New();
        sheet["A1"].Value = value;
        Assert.Equal(expected, sheet["A1"].Text);
    }

    [Theory]
    [InlineData(0.000012345)]
    [InlineData(-0.000000123456789)]
    [InlineData(99999999999.5)]
    [InlineData(1e-300)]
    public void General_text_is_what_TEXT_gives_for_General(double value)
    {
        var (workbook, sheet) = New();
        sheet["A1"].Value = value;
        Assert.Equal(workbook.Evaluate("=TEXT(S!A1,\"General\")").AsText(), sheet["A1"].Text);
    }

    [Fact]
    public void Text_of_text_boolean_error_and_empty()
    {
        var (workbook, sheet) = New();
        sheet["A1"].Value = "hello";
        sheet["A1"].NumberFormat = "0.00";
        sheet["A2"].Value = "hello";
        sheet["A2"].NumberFormat = "\"Name: \"@";
        sheet["A3"].Value = true;
        sheet["A4"].Formula = "=1/0";
        sheet["A4"].NumberFormat = "0.00";
        sheet["A5"].NumberFormat = "0.00";
        workbook.Recalculate();
        Assert.Equal("hello", sheet["A1"].Text);
        Assert.Equal("Name: hello", sheet["A2"].Text);
        Assert.Equal("TRUE", sheet["A3"].Text);
        Assert.Equal("#DIV/0!", sheet["A4"].Text);
        Assert.Equal("", sheet["A5"].Text);
        Assert.Equal("", sheet["Z9"].Text);
    }

    [Fact]
    public void Text_uses_the_workbook_culture()
    {
        var workbook = new Workbook { Culture = CultureInfo.GetCultureInfo("de-DE") };
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = 1234.5;
        sheet["A1"].NumberFormat = "#,##0.00";
        Assert.Equal("1.234,50", sheet["A1"].Text);
    }

    [Fact]
    public void A_spilled_cell_shows_its_value_with_its_own_format()
    {
        var (workbook, sheet) = New();
        sheet["A1"].Formula = "=SEQUENCE(2)/4";
        sheet["A2"].NumberFormat = "0%";
        workbook.Recalculate();
        Assert.Equal("50%", sheet["A2"].Text);
    }

    [Fact]
    public void A_number_no_date_can_show_as_hashes()
    {
        var (_, sheet) = New();
        sheet["A1"].Value = -1;
        sheet["A1"].NumberFormat = "yyyy-mm-dd";
        Assert.Equal("#######", sheet["A1"].Text);
    }
}
