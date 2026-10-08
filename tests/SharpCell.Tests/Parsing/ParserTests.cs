using SharpCell;
using SharpCell.Parsing;

namespace SharpCell.Tests.Parsing;

public class ParserTests
{
    private static string Parse(string text) =>
        TreeDump.Dump(FormulaParser.Parse(text, new CellAddress(1, 1)));

    private static FormulaParseException Fails(string text) =>
        Assert.Throws<FormulaParseException>(() => FormulaParser.Parse(text, new CellAddress(1, 1)));

    [Theory]
    [InlineData("1+2*3", "(+ 1 (* 2 3))")]
    [InlineData("1*2+3", "(+ (* 1 2) 3)")]
    [InlineData("1-2-3", "(- (- 1 2) 3)")]
    [InlineData("8/4/2", "(/ (/ 8 4) 2)")]
    [InlineData("2^3^2", "(^ (^ 2 3) 2)")]
    [InlineData("-2^2", "(^ (neg 2) 2)")]
    [InlineData("2^-2", "(^ 2 (neg 2))")]
    [InlineData("-5%", "(% (neg 5))")]
    [InlineData("5%%", "(% (% 5))")]
    [InlineData("2*50%", "(* 2 (% 50))")]
    [InlineData("--1", "(neg (neg 1))")]
    [InlineData("+1", "(pos 1)")]
    [InlineData("1+2&3", "(& (+ 1 2) 3)")]
    [InlineData("1&2=3", "(= (& 1 2) 3)")]
    [InlineData("1=2=3", "(= (= 1 2) 3)")]
    [InlineData("1<>2", "(<> 1 2)")]
    [InlineData("1<=2", "(<= 1 2)")]
    [InlineData("1>=2", "(>= 1 2)")]
    [InlineData("1<2", "(< 1 2)")]
    [InlineData("1>2", "(> 1 2)")]
    [InlineData("(1+2)*3", "(* (paren (+ 1 2)) 3)")]
    public void Operator_precedence_and_associativity(string text, string expected)
    {
        Assert.Equal(expected, Parse(text));
    }

    [Theory]
    [InlineData("1.5", "1.5")]
    [InlineData("\"a\"\"b\"", "\"a\"b\"")]
    [InlineData("TRUE", "TRUE")]
    [InlineData("#N/A", "#N/A")]
    [InlineData("A1", "(ref A1)")]
    [InlineData("'My Sheet'!$A$1:B2", "(ref My Sheet!$A$1:B2)")]
    [InlineData("Sheet1:Sheet3!A1", "(ref Sheet1:Sheet3!A1)")]
    [InlineData("Rate", "(name Rate)")]
    [InlineData("Sheet1!Rate", "(name Sheet1!Rate)")]
    [InlineData("Sheet1!#REF!+1", "(+ (referror Sheet1!) 1)")]
    [InlineData("SUM(Table1[Col])", "(fn SUM (struct Table1[Col]))")]
    public void Operands(string text, string expected)
    {
        Assert.Equal(expected, Parse(text));
    }

    [Theory]
    [InlineData("-A1:B2", "(neg (ref A1:B2))")]
    [InlineData("A1:B2:C3", "(range (ref A1:B2) (ref C3))")]
    [InlineData("A1: B2", "(ref A1:B2)")]
    [InlineData("Sheet1!A1 :B2", "(ref Sheet1!A1:B2)")]
    [InlineData("A1:Sheet1!B2", "(range (ref A1) (ref Sheet1!B2))")]
    [InlineData("A1:INDEX(B:B,2)", "(range (ref A1) (fn INDEX (ref B:B) 2))")]
    [InlineData("A1:B2 B1:C3", "(isect (ref A1:B2) (ref B1:C3))")]
    [InlineData("A1 B1 C1", "(isect (isect (ref A1) (ref B1)) (ref C1))")]
    [InlineData("SUM(A1 , B1)", "(fn SUM (ref A1) (ref B1))")]
    [InlineData("A1 -B1", "(- (ref A1) (ref B1))")]
    [InlineData("A1 + B1", "(+ (ref A1) (ref B1))")]
    [InlineData("SUM((A1,B1),C1)", "(fn SUM (paren (union (ref A1) (ref B1))) (ref C1))")]
    [InlineData("(A1,B1,C1)", "(paren (union (union (ref A1) (ref B1)) (ref C1)))")]
    [InlineData("(A1,SUM(B1,C1))", "(paren (union (ref A1) (fn SUM (ref B1) (ref C1))))")]
    [InlineData("A1#", "(spill (ref A1))")]
    [InlineData("SUM(Sheet1!A1#)", "(fn SUM (spill (ref Sheet1!A1)))")]
    [InlineData("@A1:A10", "(single (ref A1:A10))")]
    [InlineData("-@A1:A10", "(neg (single (ref A1:A10)))")]
    public void Reference_operators(string text, string expected)
    {
        Assert.Equal(expected, Parse(text));
    }

    [Fact]
    public void Union_is_only_allowed_in_parentheses()
    {
        Assert.Equal(2, Fails("A1,B1").Position);
    }

    [Theory]
    [InlineData("F()", "(fn F)")]
    [InlineData("F( )", "(fn F)")]
    [InlineData("F(,)", "(fn F _ _)")]
    [InlineData("F(1,)", "(fn F 1 _)")]
    [InlineData("F(,1)", "(fn F _ 1)")]
    [InlineData("IF(A1,,B1)", "(fn IF (ref A1) _ (ref B1))")]
    [InlineData("sum(1)", "(fn SUM 1)")]
    [InlineData("_xlfn.XLOOKUP(1,A:A,B:B)", "(fn XLOOKUP 1 (ref A:A) (ref B:B))")]
    [InlineData("_xlfn._xlws.SORT(A1:A3)", "(fn SORT (ref A1:A3))")]
    [InlineData("_xlfn.ANCHORARRAY(A1)", "(spill (ref A1))")]
    [InlineData("_xlfn.SINGLE(A1:A3)", "(single (ref A1:A3))")]
    [InlineData("_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)", "(fn LAMBDA (name x) (+ (name x) 1))")]
    [InlineData("LAMBDA(x,x+1)(2)", "(call (fn LAMBDA (name x) (+ (name x) 1)) 2)")]
    [InlineData("TRUE()", "(fn TRUE)")]
    [InlineData("_xlpm.x1", "(name _xlpm.x1)")]
    [InlineData("_xlpm.xA1+1", "(+ (name _xlpm.xA1) 1)")]
    [InlineData("_xlpm.TRUE", "(name _xlpm.TRUE)")]
    [InlineData("_xlpm.", "(name _xlpm.)")]
    [InlineData("_xlfn.(1)", "(fn _XLFN. 1)")]
    [InlineData("_xlfn.1X(1)", "(fn _XLFN.1X 1)")]
    public void Function_calls(string text, string expected)
    {
        Assert.Equal(expected, Parse(text));
    }

    [Theory]
    [InlineData("_xlfn.ANCHORARRAY(A1,B1)")]
    [InlineData("_xlfn.SINGLE()")]
    public void Special_functions_require_one_argument(string text)
    {
        Fails(text);
    }

    [Theory]
    [InlineData("{1,2;3,4}", "(array {1,2;3,4})")]
    [InlineData("{-1,\"a\";TRUE,#N/A}", "(array {-1,\"a\";TRUE,#N/A})")]
    [InlineData("{+1}", "(array {1})")]
    [InlineData("SUM({1;2})", "(fn SUM (array {1;2}))")]
    public void Array_literals(string text, string expected)
    {
        Assert.Equal(expected, Parse(text));
    }

    [Theory]
    [InlineData("{1,2;3}", 6)]
    [InlineData("{}", 1)]
    [InlineData("{A1}", 1)]
    [InlineData("{1+2}", 2)]
    [InlineData("{-\"a\"}", 2)]
    [InlineData("{{1}}", 1)]
    public void Malformed_array_literals(string text, int position)
    {
        Assert.Equal(position, Fails(text).Position);
    }

    [Fact]
    public void Leading_equals_sign_is_optional()
    {
        Assert.Equal(Parse("1+2"), Parse("=1+2"));
    }

    [Theory]
    [InlineData("", 0)]
    [InlineData("=", 1)]
    [InlineData("1+", 2)]
    [InlineData("(1", 2)]
    [InlineData("1)", 1)]
    [InlineData("SUM(1,2", 7)]
    [InlineData("1 2", 2)]
    [InlineData("*1", 0)]
    [InlineData("A1 ,B1", 3)]
    [InlineData("=+", 2)]
    public void Malformed_formulas_report_position(string text, int position)
    {
        Assert.Equal(position, Fails(text).Position);
    }

    [Theory]
    [InlineData("(", ")")]
    [InlineData("SUM(", ")")]
    [InlineData("-", "")]
    [InlineData("@", "")]
    public void Nesting_deeper_than_limit_fails(string open, string close)
    {
        var depth = FormulaLimits.MaxDepth + 1;
        var text = string.Concat(Enumerable.Repeat(open, depth)) + "1" + string.Concat(Enumerable.Repeat(close, depth));
        Fails(text);
    }

    [Theory]
    [InlineData("(", ")")]
    [InlineData("SUM(", ")")]
    [InlineData("-", "")]
    public void Nesting_at_limit_parses_on_a_512KB_stack(string open, string close)
    {
        var depth = FormulaLimits.MaxDepth;
        var text = string.Concat(Enumerable.Repeat(open, depth)) + "1" + string.Concat(Enumerable.Repeat(close, depth));
        Exception? caught = null;
        var thread = new Thread(() =>
        {
            try
            {
                FormulaParser.Parse(text, new CellAddress(1, 1));
            }
            catch (Exception ex)
            {
                caught = ex;
            }
        }, maxStackSize: 512 * 1024);
        thread.Start();
        thread.Join();
        Assert.Null(caught);
    }

    [Fact]
    public void R1C1_style()
    {
        var origin = new CellAddress(5, 1);
        var node = (FunctionNode)FormulaParser.Parse("SUM(R[-1]C:R[-1]C[2])", origin, ReferenceStyle.R1C1);
        var reference = Assert.IsType<ReferenceNode>(Assert.Single(node.Arguments));
        Assert.Equal("A4:C4", ReferenceSyntax.FormatA1(reference.Area, origin));
    }

    [Theory]
    [InlineData("SUM(", ")")]
    [InlineData("(", ")")]
    [InlineData("-", "")]
    public void Deep_nesting_on_a_small_stack_fails_cleanly(string open, string close)
    {
        var text = string.Concat(Enumerable.Repeat(open, 1000)) + "1" + string.Concat(Enumerable.Repeat(close, 1000));
        Assert.True(text.Length < FormulaLimits.MaxLength);
        Exception? caught = null;
        var thread = new Thread(() =>
        {
            try
            {
                FormulaParser.Parse(text, new CellAddress(1, 1));
            }
            catch (Exception ex)
            {
                caught = ex;
            }
        }, maxStackSize: 128 * 1024);
        thread.Start();
        thread.Join();
        Assert.IsType<FormulaParseException>(caught);
    }
}
