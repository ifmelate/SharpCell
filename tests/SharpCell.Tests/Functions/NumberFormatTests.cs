using System.Globalization;
using SharpCell;
using SharpCell.Functions;

namespace SharpCell.Tests.Functions;

public class NumberFormatTests
{
    private static string? Format(double value, string code, CultureInfo? culture = null, DateSystem system = DateSystem.Date1900)
    {
        Assert.True(NumberFormat.TryParse(code, out var format), code);
        return format!.Format(value, culture ?? CultureInfo.InvariantCulture, system);
    }

    [Theory]
    [InlineData(12, "0", "12")]
    [InlineData(12.345, "0.00", "12.35")]
    [InlineData(2.675, "0.00", "2.68")]
    [InlineData(0.5, "0", "1")]
    [InlineData(-0.5, "0", "-1")]
    [InlineData(-1.0 / 3, "0", "0")]
    [InlineData(-1.0 / 3, "0.00", "-0.33")]
    [InlineData(12345.6789, "#,##0", "12,346")]
    [InlineData(1234567, "#,##0.00", "1,234,567.00")]
    [InlineData(5, "0,000", "0,005")]
    [InlineData(12, "#.#", "12.")]
    [InlineData(0, "#.#", ".")]
    [InlineData(1.0 / 3, "#.#", ".3")]
    [InlineData(12, "000", "012")]
    [InlineData(5, "??", " 5")]
    [InlineData(1.5, "0.0?", "1.5 ")]
    [InlineData(12, ".00", "12.00")]
    [InlineData(12345.6789, "00-00-00-00", "00-01-23-46")]
    [InlineData(123456789012, "00-00-00-00", "123456-78-90-12")]
    [InlineData(-12, "00-00-00-00", "-00-00-00-12")]
    [InlineData(123456789012, "#,##,, \"Millions\"", "123,457 Millions")]
    [InlineData(1234567, "0.0,,", "1.2")]
    [InlineData(12, "0.00%", "1200.00%")]
    [InlineData(1.0 / 3, "0.00%", "33.33%")]
    [InlineData(-2.7e-18, "0.00%", "0.00%")]
    [InlineData(1.23456e19, "0", "12345600000000000000")]
    [InlineData(12, "#,#.00 \"Millions!\"", "12.00 Millions!")]
    [InlineData(12, "[Red]#.#", "12.")]
    [InlineData(12, "\\$0", "$12")]
    [InlineData(-12, "$#,##0", "-$12")]
    [InlineData(12, "0_)", "12 ")]
    [InlineData(12, "* 0", "12")]
    [InlineData(1234.5, "[$€-407] #,##0.00", "€ 1,234.50")]
    [InlineData(1234.5, "[$-409]0.0", "1234.5")]
    [InlineData(5, "\"abc\"", "abc")]
    [InlineData(-5, "\"abc\"", "-abc")]
    public void Plain_numbers(double value, string code, string expected) => Assert.Equal(expected, Format(value, code));

    [Theory]
    [InlineData(12, "#,##0_);(#,##0)", "12 ")]
    [InlineData(-12, "#,##0_);(#,##0)", "(12)")]
    [InlineData(-1.0 / 3, "#,##0_);(#,##0)", "(0)")]
    [InlineData(0, "0;-0;\"zero\"", "zero")]
    [InlineData(0, "0;-0", "0")]
    [InlineData(-5, "0;", "")]
    [InlineData(5, "[>100]\"big\";\"small\"", "small")]
    [InlineData(500, "[>100]\"big\";\"small\"", "big")]
    [InlineData(5, "[<10]0.0;[>100]0;\"mid\"", "5.0")]
    [InlineData(500, "[<10]0.0;[>100]0;\"mid\"", "500")]
    [InlineData(50, "[<10]0.0;[>100]0;\"mid\"", "mid")]
    public void Sections_and_conditions(double value, string code, string expected) => Assert.Equal(expected, Format(value, code));

    [Theory]
    [InlineData(12, "0.00E+0", "1.20E+1")]
    [InlineData(1.0 / 3, "0.00E+0", "3.33E-1")]
    [InlineData(0, "0.00E+0", "0.00E+0")]
    [InlineData(-2.7e-18, "0.00E+0", "-2.70E-18")]
    [InlineData(123456, "0.00E+00", "1.23E+05")]
    [InlineData(0.00012, "0.0E-00", "1.2E-04")]
    [InlineData(1200, "0.0E-00", "1.2E03")]
    [InlineData(9.999, "0.00E+00", "1.00E+01")]
    [InlineData(12345, "##0.0E+0", "12.3E+3")]
    [InlineData(12345, "00.00E+00", "12.35E+03")]
    public void Scientific(double value, string code, string expected) => Assert.Equal(expected, Format(value, code));

    [Theory]
    [InlineData(1.5, "# ?/?", "1 1/2")]
    [InlineData(0.75, "?/?", "3/4")]
    [InlineData(1.25, "?/?", "5/4")]
    [InlineData(3.14159, "# ??/??", "3 14/99")]
    [InlineData(1.5, "# ??/??", "1  1/2 ")]
    [InlineData(0.3, "# ?/8", " 2/8")]
    [InlineData(2.5, "# ??/100", "2 50/100")]
    [InlineData(2, "# ?/?", "2    ")]
    public void Fractions(double value, string code, string expected) => Assert.Equal(expected, Format(value, code));

    [Theory]
    [InlineData(41181, "m/d/y", "9/29/12")]
    [InlineData(41181, "dd/mm--yy", "29/09--12")]
    [InlineData(41181, "ddd--/--mmm--/++yyy", "Sat--/--Sep--/++2012")]
    [InlineData(41181, "dddd/mmmm/yyyyy", "Saturday/September/2012")]
    [InlineData(41181, "ddddd/mmmmm/yyyyy", "Saturday/S/2012")]
    [InlineData(100, "dddd d mmmm yyyy", "Monday 9 April 1900")]
    [InlineData(60, "yyyy-mm-dd", "1900-02-29")]
    [InlineData(0, "yyyy-mm-dd ddd", "1900-01-00 Sat")]
    [InlineData(2958465, "d.m.yyyy", "31.12.9999")]
    [InlineData(0.5, "h:mm", "12:00")]
    [InlineData(0.75, "h:mm AM/PM", "6:00 PM")]
    [InlineData(0.25, "hh:mm a/p", "06:00 a")]
    [InlineData(0, "h AM/PM", "12 AM")]
    [InlineData(1.5, "[h]:mm:ss", "36:00:00")]
    [InlineData(0.0423, "[mm]:ss", "60:55")]
    [InlineData(0.000011574, "[s]", "1")]
    [InlineData(0.000011574, "[ss]", "01")]
    [InlineData(0.5000058, "hh:mm:ss.00", "12:00:00.50")]
    [InlineData(0.99999999, "yyyy-mm-dd hh:mm:ss", "1900-01-01 00:00:00")]
    [InlineData(43831.5, "d-mmm-yy h:mm", "1-Jan-20 12:00")]
    [InlineData(45125, "mmm", "Jul")]
    public void Dates_and_times(double value, string code, string expected) => Assert.Equal(expected, Format(value, code));

    [Theory]
    [InlineData(-1, "m/d/y")]
    [InlineData(2958466, "m/d/y")]
    public void Dates_outside_the_calendar_fail(double value, string code) => Assert.Null(Format(value, code));

    [Fact]
    public void The_1904_system_shifts_dates_and_weekdays() =>
        Assert.Equal("Fri 1904-01-01", Format(0, "ddd yyyy-mm-dd", system: DateSystem.Date1904));

    [Theory]
    [InlineData("mm###")]
    [InlineData("d 0")]
    [InlineData("0;0;0;0;0")]
    [InlineData("\"open")]
    [InlineData("[Red")]
    [InlineData("[<abc]0")]
    public void Invalid_codes_are_rejected(string code) => Assert.False(NumberFormat.TryParse(code, out _));

    [Theory]
    [InlineData(12, "12")]
    [InlineData(-12, "-12")]
    [InlineData(1.0 / 3, "0.333333333")]
    [InlineData(-1.0 / 3, "-0.333333333")]
    [InlineData(12345.6789, "12345.6789")]
    [InlineData(123456789012, "1.23457E+11")]
    [InlineData(25000000000, "25000000000")]
    [InlineData(250000000000, "2.5E+11")]
    [InlineData(1e16, "1E+16")]
    [InlineData(1.2e-16, "1.2E-16")]
    [InlineData(-2.5e-9, "-2.5E-09")]
    [InlineData(-1e-8, "-0.00000001")]
    [InlineData(0.000012345, "0.000012345")]
    [InlineData(0.0000123456, "1.23456E-05")]
    [InlineData(99999999999.7, "1E+11")]
    [InlineData(0, "0")]
    public void General_fits_eleven_characters(double value, string expected)
    {
        Assert.Equal(expected, NumberFormat.FormatGeneral(value, CultureInfo.InvariantCulture));
        Assert.Equal(expected, Format(value, "General"));
    }

    [Theory]
    [InlineData("abc", "@", "abc")]
    [InlineData("abc", "\"<\"@\">\"", "<abc>")]
    [InlineData("abc", "0;0;0;\"x\"@", "xabc")]
    [InlineData("abc", "0;0;0;", "")]
    [InlineData("abc", "0.00", "abc")]
    public void Text_goes_through_the_text_section(string text, string code, string expected)
    {
        Assert.True(NumberFormat.TryParse(code, out var format));
        Assert.Equal(expected, format!.FormatText(text));
    }

    [Fact]
    public void A_lone_text_section_also_shows_numbers() => Assert.Equal("<5>", Format(5, "\"<\"@\">\""));

    [Fact]
    public void Output_uses_the_culture_separators_and_names()
    {
        var german = CultureInfo.GetCultureInfo("de-DE");
        Assert.Equal("1.234.567,50", Format(1234567.5, "#,##0.00", german));
        Assert.Equal("Samstag, 29. September 2012", Format(41181, "dddd, d. mmmm yyyy", german));
    }
}
