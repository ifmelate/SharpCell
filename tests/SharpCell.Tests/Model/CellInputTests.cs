using System.Globalization;

namespace SharpCell.Tests.Model;

public class CellInputTests
{
    private static Cell Typed(string input, string culture = "en-US", string? format = null)
    {
        var workbook = new Workbook { Culture = CultureInfo.GetCultureInfo(culture), Clock = new FixedClock() };
        var cell = workbook.AddSheet("S")["A1"];
        if (format is not null)
            cell.NumberFormat = format;
        cell.Input = input;
        return cell;
    }

    [Theory]
    [InlineData("42", 42.0, "General")]
    [InlineData("-1.5", -1.5, "General")]
    [InlineData("(7)", -7.0, "General")]
    [InlineData("12%", 0.12, "0%")]
    [InlineData("12.5%", 0.125, "0.00%")]
    [InlineData("1,250", 1250.0, "#,##0")]
    [InlineData("1,250.50", 1250.5, "#,##0.00")]
    [InlineData("$5", 5.0, "\"$\"#,##0")]
    [InlineData("$1,250.50", 1250.5, "\"$\"#,##0.00")]
    [InlineData("1e5", 100000.0, "0.00E+00")]
    [InlineData(" 42 ", 42.0, "General")]
    public void Numbers_are_read_and_give_a_general_cell_their_format(string input, double value, string format)
    {
        var cell = Typed(input);

        Assert.Equal(CellValue.Number(value), cell.Value);
        Assert.Equal(format, cell.NumberFormat);
    }

    [Theory]
    [InlineData("10/11/2026", 46306.0, "m/d/yyyy")]
    [InlineData("2026-10-11", 46306.0, "m/d/yyyy")]
    [InlineData("11-Oct-2026", 46306.0, "m/d/yyyy")]
    [InlineData("13:30", 0.5625, "h:mm")]
    [InlineData("1:30 PM", 0.5625, "h:mm AM/PM")]
    [InlineData("13:30:15", 0.56267361111111108, "h:mm:ss")]
    [InlineData("10/11/2026 13:30", 46306.5625, "m/d/yyyy h:mm")]
    [InlineData("3/15", 46096.0, "m/d/yyyy")]   // the current year, 2026 by the workbook clock
    public void Dates_and_times_are_read_as_serials_with_a_date_or_time_format(string input, double serial, string format)
    {
        var cell = Typed(input);

        Assert.Equal(serial, cell.Value.AsNumber(), 9);
        Assert.Equal(format, cell.NumberFormat);
    }

    [Fact]
    public void Dates_follow_the_culture()
    {
        var cell = Typed("11.10.2026", "de-DE");

        Assert.Equal(46306.0, cell.Value.AsNumber());
        Assert.Equal("dd\\.mm\\.yyyy", cell.NumberFormat);
        Assert.Equal("11.10.2026", cell.Text);
    }

    [Fact]
    public void Numbers_follow_the_culture()
    {
        var cell = Typed("1.250,5", "de-DE");

        Assert.Equal(1250.5, cell.Value.AsNumber());
        Assert.Equal("#,##0.00", cell.NumberFormat);
    }

    [Fact]
    public void A_cell_with_its_own_format_keeps_it()
    {
        var cell = Typed("12%", format: "0.000");

        Assert.Equal(0.12, cell.Value.AsNumber());
        Assert.Equal("0.000", cell.NumberFormat);
    }

    [Theory]
    [InlineData("TRUE", true)]
    [InlineData("false", false)]
    public void Logical_values_are_read(string input, bool value)
    {
        Assert.Equal(CellValue.Boolean(value), Typed(input).Value);
    }

    [Fact]
    public void Error_literals_are_read()
    {
        Assert.Equal(CellValue.Error(ErrorKind.NA), Typed("#N/A").Value);
    }

    [Theory]
    [InlineData("hello")]
    [InlineData("12 apples")]
    [InlineData("=")]
    [InlineData("  spaced  ")]
    public void Anything_else_is_text_as_typed(string input)
    {
        Assert.Equal(CellValue.Text(input), Typed(input).Value);
    }

    [Fact]
    public void An_apostrophe_keeps_the_rest_as_text()
    {
        var cell = Typed("'0123");

        Assert.Equal(CellValue.Text("0123"), cell.Value);
        Assert.Equal("General", cell.NumberFormat);
    }

    [Fact]
    public void An_equals_sign_starts_a_formula()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = 2;

        sheet["B1"].Input = "=A1*3";
        workbook.Recalculate();

        Assert.Equal("=A1*3", sheet["B1"].Formula);
        Assert.Equal(6.0, sheet["B1"].Value.AsNumber());
    }

    [Fact]
    public void A_broken_formula_throws_and_leaves_the_cell_unchanged()
    {
        var cell = Typed("5");

        Assert.Throws<FormulaParseException>(() => cell.Input = "=1+");
        Assert.Equal(5.0, cell.Value.AsNumber());
    }

    [Fact]
    public void An_empty_input_clears_the_content_but_not_the_format()
    {
        var cell = Typed("=1+1");
        cell.NumberFormat = "0.00";

        cell.Input = "";

        Assert.Null(cell.Formula);
        Assert.Equal(CellValue.Empty, cell.Value);
        Assert.Equal("0.00", cell.NumberFormat);
    }

    [Theory]
    [InlineData("42", "42")]
    [InlineData("1,250.50", "1250.5")]
    [InlineData("12.5%", "12.5%")]
    [InlineData("$5", "5")]
    [InlineData("10/11/2026", "10/11/2026")]
    [InlineData("13:30", "1:30:00 PM")]
    [InlineData("10/11/2026 13:30", "10/11/2026 1:30:00 PM")]
    [InlineData("hello", "hello")]
    [InlineData("'0123", "'0123")]
    [InlineData("'TRUE", "'TRUE")]
    [InlineData("TRUE", "TRUE")]
    [InlineData("#N/A", "#N/A")]
    [InlineData("=1+1", "=1+1")]
    [InlineData("", "")]
    public void The_edit_text_is_what_Excel_shows_to_edit_the_cell(string input, string edit)
    {
        Assert.Equal(edit, Typed(input).Input);
    }

    [Theory]
    [InlineData("1,250.50")]
    [InlineData("12.5%")]
    [InlineData("10/11/2026 13:30")]
    [InlineData("'0123")]
    [InlineData("0.1")]
    public void Typing_the_edit_text_back_keeps_the_content(string input)
    {
        var cell = Typed(input);
        var (value, format) = (cell.Value, cell.NumberFormat);

        cell.Input = cell.Input;

        Assert.Equal(value, cell.Value);
        Assert.Equal(format, cell.NumberFormat);
    }

    [Fact]
    public void A_date_shows_in_the_cultures_order()
    {
        Assert.Equal("11.10.2026", Typed("11.10.2026", "de-DE").Input);
    }

    private sealed class FixedClock : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => new(2026, 10, 11, 12, 0, 0, TimeSpan.Zero);

        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }
}
