using SharpCell;

namespace SharpCell.Tests.Values;

public class CellValueTests
{
    [Fact]
    public void Default_is_empty()
    {
        Assert.Equal(CellValueKind.Empty, default(CellValue).Kind);
        Assert.Equal(CellValue.Empty, default);
    }

    [Fact]
    public void Four_kinds_of_nothing_are_distinct()
    {
        var values = new[] { CellValue.Empty, CellValue.Missing, CellValue.Number(0), CellValue.Text("") };
        for (var i = 0; i < values.Length; i++)
            for (var j = 0; j < values.Length; j++)
                Assert.Equal(i == j, values[i].Equals(values[j]));
    }

    [Fact]
    public void Scalars_expose_their_payload()
    {
        Assert.Equal(2.5, CellValue.Number(2.5).AsNumber());
        Assert.Equal("abc", CellValue.Text("abc").AsText());
        Assert.True(CellValue.Boolean(true).AsBoolean());
        Assert.Equal(ErrorKind.Ref, CellValue.Error(ErrorKind.Ref).AsError());
    }

    [Fact]
    public void Wrong_accessor_throws()
    {
        Assert.Throws<InvalidOperationException>(() => CellValue.Text("1").AsNumber());
        Assert.Throws<InvalidOperationException>(() => CellValue.Empty.AsText());
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void Non_finite_number_becomes_NUM_error(double number)
    {
        Assert.Equal(CellValue.Error(ErrorKind.Num), CellValue.Number(number));
    }

    [Fact]
    public void Negative_zero_is_normalized()
    {
        var value = CellValue.Number(-0.0);
        Assert.False(double.IsNegative(value.AsNumber()));
        Assert.Equal(CellValue.Number(0), value);
    }

    [Fact]
    public void Text_is_compared_ordinally_for_identity()
    {
        Assert.NotEqual(CellValue.Text("a"), CellValue.Text("A"));
    }

    [Fact]
    public void Arrays_are_compared_by_content()
    {
        var a = CellValue.Array(new CellValue[,] { { 1, "x" }, { true, CellValue.Error(ErrorKind.NA) } });
        var b = CellValue.Array(new CellValue[,] { { 1, "x" }, { true, CellValue.Error(ErrorKind.NA) } });
        var c = CellValue.Array(new CellValue[,] { { 1, "x", 3 } });
        Assert.Equal(a, b);
        Assert.Equal(a.GetHashCode(), b.GetHashCode());
        Assert.NotEqual(a, c);
    }

    [Fact]
    public void Empty_array_is_rejected()
    {
        Assert.Throws<ArgumentException>(() => CellValue.Array(new CellValue[0, 3]));
    }

    [Fact]
    public void Implicit_conversions_from_primitives()
    {
        CellValue n = 2;
        CellValue s = "a";
        CellValue b = false;
        Assert.Equal(CellValue.Number(2), n);
        Assert.Equal(CellValue.Text("a"), s);
        Assert.Equal(CellValue.Boolean(false), b);
    }

    [Theory]
    [InlineData(1.5, "1.5")]
    [InlineData(-3, "-3")]
    [InlineData(0.1, "0.1")]
    public void Number_ToString_is_culture_invariant(double number, string expected)
    {
        Assert.Equal(expected, CellValue.Number(number).ToString());
    }

    [Fact]
    public void ToString_formats_like_formula_literals()
    {
        Assert.Equal("\"a\"\"b\"", CellValue.Text("a\"b").ToString());
        Assert.Equal("TRUE", CellValue.Boolean(true).ToString());
        Assert.Equal("#N/A", CellValue.Error(ErrorKind.NA).ToString());
        Assert.Equal("{1,\"x\";TRUE,#N/A}",
            CellValue.Array(new CellValue[,] { { 1, "x" }, { true, CellValue.Error(ErrorKind.NA) } }).ToString());
    }
}
