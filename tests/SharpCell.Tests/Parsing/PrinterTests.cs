using SharpCell;
using SharpCell.Parsing;

namespace SharpCell.Tests.Parsing;

public class PrinterTests
{
    private static readonly CellAddress Origin = new(3, 3);

    private static string RoundTrip(string text, ReferenceStyle style = ReferenceStyle.A1) =>
        FormulaPrinter.Print(FormulaParser.Parse(text, Origin, style), Origin, style);

    [Theory]
    // literals
    [InlineData("1")]
    [InlineData("0.5")]
    [InlineData("1E+20")]
    [InlineData("0.0000000001")]
    [InlineData("3+0.0000000000000001")]
    [InlineData("1E-30")]
    [InlineData("\"\"")]
    [InlineData("\"say \"\"hi\"\"\"")]
    [InlineData("TRUE")]
    [InlineData("#DIV/0!")]
    [InlineData("{1,2;3,4}")]
    [InlineData("{-1,\"a\";TRUE,#N/A}")]
    // operators
    [InlineData("1+2*3")]
    [InlineData("(1+2)*3")]
    [InlineData("((1))")]
    [InlineData("1-2-3")]
    [InlineData("1-(2-3)")]
    [InlineData("2^3^2")]
    [InlineData("2^(3^2)")]
    [InlineData("-2^2")]
    [InlineData("-(2^2)")]
    [InlineData("2^-2")]
    [InlineData("--1")]
    [InlineData("+A1")]
    [InlineData("-5%")]
    [InlineData("-(5%)")]
    [InlineData("50%%")]
    [InlineData("1--1")]
    [InlineData("\"a\"&B1&\"c\"")]
    [InlineData("A1=B1")]
    [InlineData("A1<>B1")]
    [InlineData("A1<=1+2")]
    [InlineData("(A1>B1)=TRUE")]
    // references
    [InlineData("A1")]
    [InlineData("$A$1:B$2")]
    [InlineData("A:A")]
    [InlineData("$1:$3")]
    [InlineData("Sheet1!A1")]
    [InlineData("'My Sheet'!A1:B2")]
    [InlineData("'It''s'!A1")]
    [InlineData("Sheet1:Sheet3!A1")]
    [InlineData("'Sheet 1:Sheet 3'!A1")]
    [InlineData("Лист1!B2")]
    [InlineData("'2020'!A1")]
    [InlineData("'A1'!B2")]
    [InlineData("'R1C1'!B2")]
    [InlineData("Sheet1!#REF!")]
    [InlineData("Rate*2")]
    [InlineData("Sheet1!Rate")]
    [InlineData("Table1[[#This Row],[Col]]")]
    [InlineData("A1:B2 B1:C3")]
    [InlineData("SUM((A1,B1),C1)")]
    [InlineData("SUM((A1,B1,C1))")]
    [InlineData("A1:INDEX(B:B,2)")]
    [InlineData("A1:B2:C3")]
    [InlineData("1 :1:3")]
    [InlineData("Rate :Other!B2")]
    [InlineData("A1#")]
    [InlineData("SUM(Sheet1!A1#)")]
    [InlineData("@A1:A10")]
    [InlineData("-@A1:A10")]
    // functions
    [InlineData("SUM(A1:A10)")]
    [InlineData("F()")]
    [InlineData("F(,)")]
    [InlineData("IF(A1,,B1)")]
    [InlineData("IF(A1>0,\"pos\",IF(A1<0,\"neg\",\"zero\"))")]
    [InlineData("IFERROR(VLOOKUP($A2,Data!$A:$D,4,FALSE),0)")]
    [InlineData("SUMPRODUCT((A1:A10>0)*(B1:B10))")]
    [InlineData("INDEX(B:B,MATCH(MAX(C:C),C:C,0))")]
    [InlineData("LET(x,A1*2,y,x+1,x*y)")]
    [InlineData("LAMBDA(x,y,x+y)(1,2)")]
    [InlineData("MAP(A1:A3,LAMBDA(v,v*2))")]
    [InlineData("REDUCE(0,A1:A10,LAMBDA(acc,v,acc+v))")]
    [InlineData("XLOOKUP(A1,B:B,C:C,\"none\")")]
    [InlineData("SORT(FILTER(A1:B10,B1:B10>5),2,-1)")]
    [InlineData("SUMIFS(C:C,A:A,\">=\"&DATE(2026,1,1),B:B,\"<>x\")")]
    [InlineData("TEXT(NOW(),\"yyyy-mm-dd\")&\" \"&Sheet2!A1")]
    [InlineData("ROUND(-(A1+B1)/2,2)")]
    [InlineData("SUM({1,2,3}*2)")]
    public void Canonical_formulas_round_trip(string text)
    {
        Assert.Equal(text, RoundTrip(text));
    }

    [Theory]
    [InlineData("=1+2", "1+2")]
    [InlineData("sum( a1 , 2 )", "SUM(A1,2)")]
    [InlineData("1 + 2", "1+2")]
    [InlineData(".5", "0.5")]
    [InlineData("1e3", "1000")]
    [InlineData("_xlfn.XLOOKUP(1,A:A,B:B)", "XLOOKUP(1,A:A,B:B)")]
    [InlineData("_xlfn.ANCHORARRAY(A1)", "A1#")]
    [InlineData("_xlfn.SINGLE(A1:A3)", "@A1:A3")]
    [InlineData("_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)", "LAMBDA(x,x+1)")]
    [InlineData("_xlfn.LET(_xlpm.area,_xlfn.LAMBDA(_xlpm.x,_xlpm.y,_xlpm.x*_xlpm.y),_xlpm.area(D1,E1))", "LET(area,LAMBDA(x,y,x*y),AREA(D1,E1))")]
    [InlineData("1E-10", "0.0000000001")]
    [InlineData("'Sheet1'!A1", "Sheet1!A1")]
    [InlineData("{+1}", "{1}")]
    [InlineData("1: 1:3", "1 :1:3")]
    [InlineData("A1 : B2", "A1:B2")]
    public void Non_canonical_input_prints_canonically(string text, string expected)
    {
        Assert.Equal(expected, RoundTrip(text));
    }

    [Theory]
    [InlineData("SUM(R[-1]C:R[1]C[2])")]
    [InlineData("R1C1+RC[-1]")]
    [InlineData("SUM(R2:R[1])")]
    [InlineData("Sheet1!C[1]")]
    public void R1C1_round_trips(string text)
    {
        Assert.Equal(text, RoundTrip(text, ReferenceStyle.R1C1));
    }

    [Fact]
    public void Prints_in_either_reference_style()
    {
        var node = FormulaParser.Parse("SUM(A1:$B$2)", Origin);
        Assert.Equal("SUM(R[-2]C[-2]:R2C2)", FormulaPrinter.Print(node, Origin, ReferenceStyle.R1C1));
    }

    [Fact]
    public void Relative_references_follow_the_origin()
    {
        var node = FormulaParser.Parse("A1+$A$1", new CellAddress(2, 2));
        Assert.Equal("A4+$A$1", FormulaPrinter.Print(node, new CellAddress(5, 2)));
    }

    [Fact]
    public void Long_operator_chain_prints_without_deep_recursion()
    {
        var text = "A1" + string.Concat(Enumerable.Repeat("+A1", 2000));
        Assert.True(text.Length < FormulaLimits.MaxLength);
        string? printed = null;
        var thread = new Thread(() => printed = RoundTrip(text), maxStackSize: 256 * 1024);
        thread.Start();
        thread.Join();
        Assert.Equal(text, printed);
    }

    private static readonly NumberNode One = new(1);
    private static readonly NumberNode Two = new(2);
    private static readonly ReferenceNode A1 = new(null, AreaRef.Cell(new CellRef(AxisRef.Absolute(1), AxisRef.Absolute(1))));
    private static readonly ReferenceNode B1 = new(null, AreaRef.Cell(new CellRef(AxisRef.Absolute(1), AxisRef.Absolute(2))));

    private static readonly (FormulaNode Node, string Expected)[] BuiltTrees =
    [
        (new BinaryNode(BinaryOperator.Multiply, new BinaryNode(BinaryOperator.Add, One, Two), Two), "(1+2)*2"),
        (new BinaryNode(BinaryOperator.Subtract, One, new BinaryNode(BinaryOperator.Subtract, One, Two)), "1-(1-2)"),
        (new BinaryNode(BinaryOperator.Power, Two, new BinaryNode(BinaryOperator.Power, Two, Two)), "2^(2^2)"),
        (new UnaryNode(UnaryOperator.Negate, new UnaryNode(UnaryOperator.Percent, One)), "-(1%)"),
        (new BinaryNode(BinaryOperator.Power, new NumberNode(-2), Two), "-2^2"),
        (new BinaryNode(BinaryOperator.Union, A1, B1), "($A$1,$B$1)"),
        (new FunctionNode("SUM", [new BinaryNode(BinaryOperator.Union, A1, B1)]), "SUM(($A$1,$B$1))"),
        (new BinaryNode(BinaryOperator.Intersect, A1, new UnaryNode(UnaryOperator.Negate, B1)), "$A$1 (-$B$1)"),
        (new BinaryNode(BinaryOperator.Intersect, A1, One), "$A$1 (1)"),
        (new BinaryNode(BinaryOperator.Range, A1, new BinaryNode(BinaryOperator.Add, B1, One)), "$A$1:($B$1+1)"),
        (new SpillNode(new BinaryNode(BinaryOperator.Range, A1, B1)), "($A$1:$B$1)#"),
        (new ImplicitIntersectionNode(new BinaryNode(BinaryOperator.Intersect, A1, B1)), "@($A$1 $B$1)"),
        (new CallNode(new NameNode(null, "f"), [One]), "(f)(1)"),
    ];

    public static TheoryData<int> BuiltTreeCases => new(Enumerable.Range(0, BuiltTrees.Length));

    [Theory]
    [MemberData(nameof(BuiltTreeCases))]
    public void Built_trees_get_only_the_parentheses_they_need(int index)
    {
        var (node, expected) = BuiltTrees[index];
        var printed = FormulaPrinter.Print(node, Origin);
        Assert.Equal(expected, printed);
        Assert.Equal(printed, RoundTrip(printed));
    }
}
