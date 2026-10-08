using SharpCell;
using SharpCell.Parsing;

namespace SharpCell.Tests.Parsing;

public class LexerTests
{
    private static readonly CellAddress Origin = new(1, 1);

    private static List<Token> Lex(string text, ReferenceStyle style = ReferenceStyle.A1) =>
        Lexer.Tokenize(text, Origin, style);

    private static Token Single(string text, ReferenceStyle style = ReferenceStyle.A1)
    {
        var tokens = Lex(text, style);
        Assert.Equal(2, tokens.Count);
        Assert.Equal(TokenKind.End, tokens[1].Kind);
        return tokens[0];
    }

    private static TokenKind[] Kinds(string text, ReferenceStyle style = ReferenceStyle.A1) =>
        Lex(text, style).Select(t => t.Kind).SkipLast(1).ToArray();

    private static FormulaParseException Fails(string text, ReferenceStyle style = ReferenceStyle.A1) =>
        Assert.Throws<FormulaParseException>(() => Lex(text, style));

    [Theory]
    [InlineData("1", 1)]
    [InlineData("1.5", 1.5)]
    [InlineData(".5", 0.5)]
    [InlineData("1.", 1)]
    [InlineData("1e3", 1000)]
    [InlineData("1E-3", 0.001)]
    [InlineData("1.5e+2", 150)]
    public void Numbers(string text, double expected)
    {
        var token = Single(text);
        Assert.Equal(TokenKind.Number, token.Kind);
        Assert.Equal(expected, token.Number);
    }

    [Theory]
    [InlineData("1e")]
    [InlineData("1e+")]
    [InlineData("1e309")]
    [InlineData("1A")]
    [InlineData("1.2.3")]
    public void Malformed_numbers_fail_at_their_start(string text)
    {
        Assert.Equal(0, Fails(text).Position);
    }

    [Theory]
    [InlineData("\"abc\"", "abc")]
    [InlineData("\"\"", "")]
    [InlineData("\"a\"\"b\"", "a\"b")]
    [InlineData("\"Привет, мир\"", "Привет, мир")]
    public void Strings(string text, string expected)
    {
        var token = Single(text);
        Assert.Equal(TokenKind.String, token.Kind);
        Assert.Equal(expected, token.Text);
    }

    [Fact]
    public void Unterminated_string_reports_its_position()
    {
        Assert.Equal(2, Fails("1+\"abc").Position);
    }

    [Theory]
    [InlineData("TRUE", true)]
    [InlineData("false", false)]
    public void Booleans(string text, bool expected)
    {
        var token = Single(text);
        Assert.Equal(TokenKind.Boolean, token.Kind);
        Assert.Equal(expected, token.Boolean);
    }

    [Fact]
    public void Boolean_followed_by_parenthesis_is_a_function()
    {
        Assert.Equal([TokenKind.Function, TokenKind.OpenParen, TokenKind.CloseParen], Kinds("TRUE()"));
    }

    [Theory]
    [InlineData("#DIV/0!", ErrorKind.Div0)]
    [InlineData("#n/a", ErrorKind.NA)]
    [InlineData("#SPILL!", ErrorKind.Spill)]
    public void Errors(string text, ErrorKind expected)
    {
        var token = Single(text);
        Assert.Equal(TokenKind.Error, token.Kind);
        Assert.Equal(expected, token.Error);
    }

    [Fact]
    public void Operators_and_punctuation()
    {
        Assert.Equal(
        [
            TokenKind.NotEqual, TokenKind.LessEqual, TokenKind.GreaterEqual, TokenKind.Less, TokenKind.Greater,
            TokenKind.Equal, TokenKind.Plus, TokenKind.Minus, TokenKind.Star, TokenKind.Slash, TokenKind.Caret,
            TokenKind.Ampersand, TokenKind.Percent, TokenKind.OpenParen, TokenKind.CloseParen, TokenKind.Comma,
            TokenKind.Semicolon, TokenKind.OpenBrace, TokenKind.CloseBrace, TokenKind.At, TokenKind.Colon,
        ], Kinds("<> <= >= < > = + - * / ^ & % ( ) , ; { } @ :"));
    }

    [Theory]
    [InlineData("A1", "A1")]
    [InlineData("$A$1:b2", "$A$1:B2")]
    [InlineData("A:A", "A:A")]
    [InlineData("1:3", "1:3")]
    [InlineData("$1:$3", "$1:$3")]
    [InlineData("XFD1048576", "XFD1048576")]
    public void A1_references_are_single_tokens(string text, string expected)
    {
        var token = Single(text);
        Assert.Equal(TokenKind.Reference, token.Kind);
        Assert.Null(token.Sheet);
        Assert.Equal(expected, ReferenceSyntax.FormatA1(token.Area, Origin));
    }

    [Fact]
    public void Incomplete_range_leaves_colon_as_operator()
    {
        Assert.Equal([TokenKind.Reference, TokenKind.Colon, TokenKind.Name], Kinds("A1:B"));
    }

    [Fact]
    public void Function_names_keep_prefixes()
    {
        var tokens = Lex("_xlfn.XLOOKUP(1)");
        Assert.Equal(TokenKind.Function, tokens[0].Kind);
        Assert.Equal("_xlfn.XLOOKUP", tokens[0].Text);
        Assert.Equal(TokenKind.OpenParen, tokens[1].Kind);
    }

    [Fact]
    public void Cell_like_function_name_is_a_function()
    {
        Assert.Equal(TokenKind.Function, Lex("LOG10(100)")[0].Kind);
    }

    [Theory]
    [InlineData("MyName")]
    [InlineData("_total")]
    [InlineData("XYZ1")]
    [InlineData("A1B")]
    [InlineData("Налог")]
    [InlineData("rate.2026")]
    public void Names(string text)
    {
        var token = Single(text);
        Assert.Equal(TokenKind.Name, token.Kind);
        Assert.Equal(text, token.Text);
    }

    [Theory]
    [InlineData("Sheet1!A1", "Sheet1", null, "A1")]
    [InlineData("'My Sheet'!A1:B2", "My Sheet", null, "A1:B2")]
    [InlineData("'It''s'!A1", "It's", null, "A1")]
    [InlineData("Sheet1:Sheet3!A1", "Sheet1", "Sheet3", "A1")]
    [InlineData("'Sheet 1:Sheet 3'!$B$2", "Sheet 1", "Sheet 3", "$B$2")]
    [InlineData("Лист1!C:C", "Лист1", null, "C:C")]
    [InlineData("'2020'!1:1", "2020", null, "1:1")]
    public void Sheet_prefixed_references(string text, string first, string? last, string area)
    {
        var token = Single(text);
        Assert.Equal(TokenKind.Reference, token.Kind);
        Assert.Equal(new SheetPrefix(first, last), token.Sheet);
        Assert.Equal(area, ReferenceSyntax.FormatA1(token.Area, Origin));
    }

    [Fact]
    public void Sheet_prefixed_name()
    {
        var token = Single("Sheet1!Rate");
        Assert.Equal(TokenKind.Name, token.Kind);
        Assert.Equal("Rate", token.Text);
        Assert.Equal(new SheetPrefix("Sheet1"), token.Sheet);
    }

    [Fact]
    public void Sheet_prefixed_REF_error()
    {
        var token = Single("Sheet1!#REF!");
        Assert.Equal(TokenKind.Error, token.Kind);
        Assert.Equal(ErrorKind.Ref, token.Error);
        Assert.Equal(new SheetPrefix("Sheet1"), token.Sheet);
    }

    [Theory]
    [InlineData("'abc", 0)]
    [InlineData("'abc'A1", 0)]
    [InlineData("1+Sheet1!", 2)]
    [InlineData("Sheet1!1", 0)]
    [InlineData("[1]Sheet1!A1", 0)]
    [InlineData("'[1]Sheet 1'!A1", 0)]
    [InlineData("1+`", 2)]
    [InlineData("A1!", 0)]
    [InlineData("$A", 0)]
    public void Malformed_input_reports_position(string text, int position)
    {
        Assert.Equal(position, Fails(text).Position);
    }

    [Fact]
    public void Whitespace_is_recorded_on_the_next_token()
    {
        var tokens = Lex("A1 B1\n+ C1");
        Assert.False(tokens[0].SpaceBefore);
        Assert.True(tokens[1].SpaceBefore);
        Assert.True(tokens[2].SpaceBefore);
        Assert.True(tokens[3].SpaceBefore);
    }

    [Fact]
    public void Hash_after_reference_is_spill_operator()
    {
        Assert.Equal([TokenKind.Reference, TokenKind.Hash], Kinds("A1#"));
    }

    [Theory]
    [InlineData("Table1[Col]")]
    [InlineData("Table1[[#This Row],[Col]]")]
    [InlineData("[@Col]")]
    [InlineData("T[a']b]")]
    public void Structured_references_are_kept_raw(string text)
    {
        var token = Single(text);
        Assert.Equal(TokenKind.StructuredReference, token.Kind);
        Assert.Equal(text, token.Text);
    }

    [Fact]
    public void Unbalanced_structured_reference_fails()
    {
        Assert.Equal(0, Fails("T[a").Position);
    }

    [Theory]
    [InlineData("R[-1]C", "A1048576")]
    [InlineData("RC:R[1]C[1]", "A1:B2")]
    [InlineData("R1C1", "$A$1")]
    [InlineData("r2", "$2:$2")]
    [InlineData("C", "A:A")]
    public void R1C1_references(string text, string a1)
    {
        var token = Single(text, ReferenceStyle.R1C1);
        Assert.Equal(TokenKind.Reference, token.Kind);
        Assert.Equal(a1, ReferenceSyntax.FormatA1(token.Area, Origin));
    }

    [Fact]
    public void R1C1_mode_still_lexes_functions_and_names()
    {
        Assert.Equal([TokenKind.Function, TokenKind.OpenParen, TokenKind.Reference, TokenKind.CloseParen],
            Kinds("ROUND(Sheet1!RC)", ReferenceStyle.R1C1));
        Assert.Equal(TokenKind.Name, Single("A1", ReferenceStyle.R1C1).Kind);
    }

    [Fact]
    public void Formula_longer_than_limit_fails()
    {
        var text = "1" + string.Concat(Enumerable.Repeat("+1", FormulaLimits.MaxLength / 2));
        Assert.Throws<FormulaParseException>(() => Lex(text));
    }

    [Fact]
    public void Token_positions_point_into_the_text()
    {
        var tokens = Lex("SUM( A1 , 2)");
        Assert.Equal((0, 3), (tokens[0].Start, tokens[0].Length));
        Assert.Equal((5, 2), (tokens[2].Start, tokens[2].Length));
        Assert.Equal((10, 1), (tokens[4].Start, tokens[4].Length));
    }
}
