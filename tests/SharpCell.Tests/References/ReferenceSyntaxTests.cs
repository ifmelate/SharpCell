using SharpCell;

namespace SharpCell.Tests.References;

public class ReferenceSyntaxTests
{
    private static readonly CellAddress C3 = new(3, 3);

    [Fact]
    public void A1_relative_parts_are_stored_as_offsets_from_origin()
    {
        Assert.True(ReferenceSyntax.TryParseA1Area("A1", C3, out var area));
        Assert.Equal(AreaRef.Cell(new CellRef(AxisRef.Relative(-2), AxisRef.Relative(-2))), area);
    }

    [Fact]
    public void A1_dollar_marks_absolute_parts()
    {
        Assert.True(ReferenceSyntax.TryParseA1Area("$A1", C3, out var area));
        Assert.Equal(AreaRef.Cell(new CellRef(AxisRef.Relative(-2), AxisRef.Absolute(1))), area);
        Assert.True(ReferenceSyntax.TryParseA1Area("a$1", C3, out area));
        Assert.Equal(AreaRef.Cell(new CellRef(AxisRef.Absolute(1), AxisRef.Relative(-2))), area);
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("$B$2")]
    [InlineData("XFD1048576")]
    [InlineData("A1:B2")]
    [InlineData("$A$1:B$2")]
    [InlineData("A:A")]
    [InlineData("$A:$C")]
    [InlineData("1:1")]
    [InlineData("$2:10")]
    [InlineData("B2:A1")]
    public void A1_round_trips(string text)
    {
        Assert.True(ReferenceSyntax.TryParseA1Area(text, C3, out var area));
        Assert.Equal(text, ReferenceSyntax.FormatA1(area, C3));
    }

    [Fact]
    public void A1_is_normalized_to_upper_case()
    {
        Assert.True(ReferenceSyntax.TryParseA1Area("a1:$b$2", C3, out var area));
        Assert.Equal("A1:$B$2", ReferenceSyntax.FormatA1(area, C3));
    }

    [Theory]
    [InlineData("")]
    [InlineData("A")]
    [InlineData("1")]
    [InlineData("A0")]
    [InlineData("XFE1")]
    [InlineData("A1048577")]
    [InlineData("A1:B")]
    [InlineData("A:1")]
    [InlineData("A1:B2:C3")]
    [InlineData("$$A1")]
    [InlineData("A1$")]
    [InlineData("A:")]
    public void Invalid_A1_is_rejected(string text)
    {
        Assert.False(ReferenceSyntax.TryParseA1Area(text, C3, out _));
    }

    [Fact]
    public void Relative_reference_moves_with_origin()
    {
        Assert.True(ReferenceSyntax.TryParseA1Area("A1:$B$1", new CellAddress(2, 2), out var area));
        Assert.Equal("A4:$B$1", ReferenceSyntax.FormatA1(area, new CellAddress(5, 2)));
    }

    [Theory]
    [InlineData("R1C1", "$A$1")]
    [InlineData("R[-1]C", "C2")]
    [InlineData("RC[2]", "E3")]
    [InlineData("RC", "C3")]
    [InlineData("R1C1:R[1]C[1]", "$A$1:D4")]
    [InlineData("R2", "$2:$2")]
    [InlineData("R[1]:R[2]", "4:5")]
    [InlineData("C1:C[1]", "$A:D")]
    [InlineData("C", "C:C")]
    public void R1C1_parses_to_the_same_reference_as_A1(string r1c1, string a1)
    {
        Assert.True(ReferenceSyntax.TryParseR1C1Area(r1c1, out var area));
        Assert.Equal(a1, ReferenceSyntax.FormatA1(area, C3));
    }

    [Theory]
    [InlineData("R1C1")]
    [InlineData("R[-1]C")]
    [InlineData("RC[2]")]
    [InlineData("RC")]
    [InlineData("R1C1:R[1]C[1]")]
    [InlineData("R2")]
    [InlineData("R[1]:R[2]")]
    [InlineData("C1:C[1]")]
    [InlineData("R1048576C16384")]
    public void R1C1_round_trips(string text)
    {
        Assert.True(ReferenceSyntax.TryParseR1C1Area(text, out var area));
        Assert.Equal(text, ReferenceSyntax.FormatR1C1(area));
    }

    [Fact]
    public void A1_converts_to_R1C1()
    {
        Assert.True(ReferenceSyntax.TryParseA1Area("B2:$D$4", C3, out var area));
        Assert.Equal("R[-1]C[-1]:R4C4", ReferenceSyntax.FormatR1C1(area));
        Assert.True(ReferenceSyntax.TryParseA1Area("C3", C3, out area));
        Assert.Equal("RC", ReferenceSyntax.FormatR1C1(area));
    }

    [Theory]
    [InlineData("")]
    [InlineData("R0C1")]
    [InlineData("R1C16385")]
    [InlineData("R1048577C1")]
    [InlineData("R[1048576]C")]
    [InlineData("R[]C")]
    [InlineData("R[1C")]
    [InlineData("R1C1:C1")]
    [InlineData("CR")]
    [InlineData("R1C1x")]
    public void Invalid_R1C1_is_rejected(string text)
    {
        Assert.False(ReferenceSyntax.TryParseR1C1Area(text, out _));
    }
}
