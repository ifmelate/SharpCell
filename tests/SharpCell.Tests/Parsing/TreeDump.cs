using System.Globalization;
using SharpCell;
using SharpCell.Parsing;

namespace SharpCell.Tests.Parsing;

/// <summary>Renders a formula tree as an S-expression so parser tests can assert on structure.</summary>
internal static class TreeDump
{
    private static readonly CellAddress Origin = new(1, 1);

    public static string Dump(FormulaNode node) => node switch
    {
        NumberNode n => n.Value.ToString("R", CultureInfo.InvariantCulture),
        TextNode t => "\"" + t.Value + "\"",
        BooleanNode b => b.Value ? "TRUE" : "FALSE",
        ErrorNode e => e.Error.ToText(),
        MissingNode => "_",
        ReferenceNode r => $"(ref {Sheet(r.Sheet)}{ReferenceSyntax.FormatA1(r.Area, Origin)})",
        RefErrorNode r => $"(referror {Sheet(r.Sheet)})",
        NameNode n => $"(name {Sheet(n.Sheet)}{n.Name})",
        StructuredReferenceNode s => $"(struct {s.Text})",
        UnaryNode u => $"({UnaryText(u.Operator)} {Dump(u.Operand)})",
        BinaryNode b => $"({BinaryText(b.Operator)} {Dump(b.Left)} {Dump(b.Right)})",
        FunctionNode f => $"(fn {f.Name}{Args(f.Arguments)})",
        CallNode c => $"(call {Dump(c.Callee)}{Args(c.Arguments)})",
        ParenthesesNode p => $"(paren {Dump(p.Inner)})",
        ArrayNode a => $"(array {CellValue.Array(a.Values)})",
        SpillNode s => $"(spill {Dump(s.Operand)})",
        ImplicitIntersectionNode i => $"(single {Dump(i.Operand)})",
        _ => throw new ArgumentException(node.GetType().Name),
    };

    private static string Args(IReadOnlyList<FormulaNode> args) => string.Concat(args.Select(a => " " + Dump(a)));

    private static string Sheet(SheetPrefix? sheet) =>
        sheet is null ? "" : sheet.Last is null ? sheet.First + "!" : $"{sheet.First}:{sheet.Last}!";

    private static string UnaryText(UnaryOperator op) => op switch
    {
        UnaryOperator.Negate => "neg",
        UnaryOperator.Plus => "pos",
        _ => "%",
    };

    private static string BinaryText(BinaryOperator op) => op switch
    {
        BinaryOperator.Range => "range",
        BinaryOperator.Intersect => "isect",
        BinaryOperator.Union => "union",
        BinaryOperator.Power => "^",
        BinaryOperator.Multiply => "*",
        BinaryOperator.Divide => "/",
        BinaryOperator.Add => "+",
        BinaryOperator.Subtract => "-",
        BinaryOperator.Concat => "&",
        BinaryOperator.Equal => "=",
        BinaryOperator.NotEqual => "<>",
        BinaryOperator.Less => "<",
        BinaryOperator.LessEqual => "<=",
        BinaryOperator.Greater => ">",
        _ => ">=",
    };
}
