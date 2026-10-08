using SharpCell;

namespace SharpCell.Tests.Values;

public class ErrorKindTests
{
    [Theory]
    [InlineData(ErrorKind.Null, "#NULL!")]
    [InlineData(ErrorKind.Div0, "#DIV/0!")]
    [InlineData(ErrorKind.Value, "#VALUE!")]
    [InlineData(ErrorKind.Ref, "#REF!")]
    [InlineData(ErrorKind.Name, "#NAME?")]
    [InlineData(ErrorKind.Num, "#NUM!")]
    [InlineData(ErrorKind.NA, "#N/A")]
    [InlineData(ErrorKind.Spill, "#SPILL!")]
    [InlineData(ErrorKind.Calc, "#CALC!")]
    public void Text_round_trips(ErrorKind kind, string text)
    {
        Assert.Equal(text, kind.ToText());
        Assert.True(ErrorKinds.TryParse(text, out var parsed));
        Assert.Equal(kind, parsed);
    }

    [Fact]
    public void Parse_ignores_case()
    {
        Assert.True(ErrorKinds.TryParse("#div/0!", out var kind));
        Assert.Equal(ErrorKind.Div0, kind);
    }

    [Theory]
    [InlineData("")]
    [InlineData("#DIV/0")]
    [InlineData("#FOO!")]
    [InlineData("DIV/0!")]
    public void Parse_rejects_unknown_text(string text)
    {
        Assert.False(ErrorKinds.TryParse(text, out _));
    }

    [Fact]
    public void Numeric_codes_match_ERROR_TYPE()
    {
        Assert.Equal(1, (int)ErrorKind.Null);
        Assert.Equal(7, (int)ErrorKind.NA);
        Assert.Equal(9, (int)ErrorKind.Spill);
        Assert.Equal(14, (int)ErrorKind.Calc);
    }
}
