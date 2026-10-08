using System.Globalization;
using SharpCell;

namespace SharpCell.Tests.Values;

public class CoercionTests
{
    private static readonly CultureInfo EnUs = CultureInfo.GetCultureInfo("en-US");
    private static readonly CultureInfo RuRu = CultureInfo.GetCultureInfo("ru-RU");

    [Theory]
    [InlineData("1", 1)]
    [InlineData(" 2.5 ", 2.5)]
    [InlineData("1,234.5", 1234.5)]
    [InlineData("1,234,567", 1234567)]
    [InlineData("1e3", 1000)]
    [InlineData("1.5E-2", 0.015)]
    [InlineData("+7", 7)]
    [InlineData("-.5", -0.5)]
    [InlineData("5.", 5)]
    [InlineData("50%", 0.5)]
    [InlineData("$5", 5)]
    [InlineData("-$5", -5)]
    [InlineData("(5)", -5)]
    public void Text_to_number_en_US(string text, double expected)
    {
        Assert.Equal(CellValue.Number(expected), Coercion.ToNumber(text, EnUs));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("abc")]
    [InlineData("TRUE")]
    [InlineData("1,23")]
    [InlineData("1.2.3")]
    [InlineData("1e")]
    [InlineData("1e309")]
    [InlineData("--1")]
    [InlineData("(-5)")]
    [InlineData("1 2")]
    public void Unparsable_text_is_VALUE_error(string text)
    {
        Assert.Equal(CellValue.Error(ErrorKind.Value), Coercion.ToNumber(text, EnUs));
    }

    [Theory]
    [InlineData("1,5", 1.5)]
    [InlineData("1 234,5", 1234.5)]
    [InlineData("1 234,5", 1234.5)]
    public void Text_to_number_uses_workbook_culture(string text, double expected)
    {
        Assert.Equal(CellValue.Number(expected), Coercion.ToNumber(text, RuRu));
    }

    [Fact]
    public void Non_text_to_number()
    {
        Assert.Equal(CellValue.Number(0), Coercion.ToNumber(CellValue.Empty, EnUs));
        Assert.Equal(CellValue.Number(0), Coercion.ToNumber(CellValue.Missing, EnUs));
        Assert.Equal(CellValue.Number(1), Coercion.ToNumber(true, EnUs));
        Assert.Equal(CellValue.Number(0), Coercion.ToNumber(false, EnUs));
        Assert.Equal(CellValue.Error(ErrorKind.NA), Coercion.ToNumber(CellValue.Error(ErrorKind.NA), EnUs));
        Assert.Equal(CellValue.Error(ErrorKind.Value), Coercion.ToNumber(CellValue.Array(new CellValue[,] { { 1 } }), EnUs));
    }

    [Theory]
    [InlineData(100, "100")]
    [InlineData(-1.5, "-1.5")]
    [InlineData(1.0 / 3, "0.333333333333333")]
    [InlineData(0.1 + 0.2, "0.3")]
    [InlineData(999999999999999, "999999999999999")]
    [InlineData(1e15, "1E+15")]
    [InlineData(1e20, "1E+20")]
    [InlineData(123456789012345678, "1.23456789012346E+17")]
    [InlineData(0.0001, "0.0001")]
    [InlineData(1e-9, "0.000000001")]
    [InlineData(1e-10, "1E-10")]
    [InlineData(1.5e-12, "1.5E-12")]
    [InlineData(9.9999999999999999, "10")]
    public void Number_to_text_uses_15_significant_digits(double number, string expected)
    {
        Assert.Equal(CellValue.Text(expected), Coercion.ToText(number, EnUs));
    }

    [Fact]
    public void Number_to_text_uses_workbook_decimal_separator()
    {
        Assert.Equal(CellValue.Text("0,5"), Coercion.ToText(0.5, RuRu));
    }

    [Fact]
    public void Non_number_to_text()
    {
        Assert.Equal(CellValue.Text(""), Coercion.ToText(CellValue.Empty, EnUs));
        Assert.Equal(CellValue.Text(""), Coercion.ToText(CellValue.Missing, EnUs));
        Assert.Equal(CellValue.Text("TRUE"), Coercion.ToText(true, EnUs));
        Assert.Equal(CellValue.Text("abc"), Coercion.ToText("abc", EnUs));
        Assert.Equal(CellValue.Error(ErrorKind.Ref), Coercion.ToText(CellValue.Error(ErrorKind.Ref), EnUs));
    }

    [Fact]
    public void To_boolean()
    {
        Assert.Equal(CellValue.True, Coercion.ToBoolean(2));
        Assert.Equal(CellValue.False, Coercion.ToBoolean(0));
        Assert.Equal(CellValue.False, Coercion.ToBoolean(CellValue.Empty));
        Assert.Equal(CellValue.False, Coercion.ToBoolean(CellValue.Missing));
        Assert.Equal(CellValue.True, Coercion.ToBoolean("true"));
        Assert.Equal(CellValue.False, Coercion.ToBoolean("FALSE"));
        Assert.Equal(CellValue.Error(ErrorKind.Value), Coercion.ToBoolean("yes"));
        Assert.Equal(CellValue.Error(ErrorKind.Value), Coercion.ToBoolean("1"));
        Assert.Equal(CellValue.Error(ErrorKind.Div0), Coercion.ToBoolean(CellValue.Error(ErrorKind.Div0)));
    }
}
