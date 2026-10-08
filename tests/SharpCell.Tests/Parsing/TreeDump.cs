using System.Globalization;
using SharpCell;
using SharpCell.Parsing;

namespace SharpCell.Tests.Parsing;

/// <summary>Renders a formula tree as an S-expression so parser tests can assert on structure.</summary>
internal static class TreeDump
{
    private static readonly CellAddress Origin = new(1, 1);

    public static string Dump(FormulaNode node) => Dump(node, ignoreParentheses: false);

    /// <param name="ignoreParentheses">Dump explicit parentheses as their content, for comparing meaning.</param>
    public static string Dump(FormulaNode node, bool ignoreParentheses) => node switch
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
        UnaryNode u => $"({UnaryText(u.Operator)} {Dump(u.Operand, ignoreParentheses)})",
        BinaryNode b => DumpBinary(b, ignoreParentheses),
        FunctionNode f => $"(fn {f.Name}{Args(f.Arguments, ignoreParentheses)})",
        CallNode c => $"(call {Dump(c.Callee, ignoreParentheses)}{Args(c.Arguments, ignoreParentheses)})",
        ParenthesesNode p => ignoreParentheses ? Dump(p.Inner, true) : $"(paren {Dump(p.Inner)})",
        ArrayNode a => $"(array {CellValue.Array(a.Values)})",
        SpillNode s => $"(spill {Dump(s.Operand, ignoreParentheses)})",
        ImplicitIntersectionNode i => $"(single {Dump(i.Operand, ignoreParentheses)})",
        _ => throw new ArgumentException(node.GetType().Name),
    };

    // Walks the left spine in a loop: parsed operator chains are thousands of nodes tall.
    private static string DumpBinary(BinaryNode node, bool ignoreParentheses)
    {
        var spine = new List<BinaryNode>();
        FormulaNode current = node;
        while (current is BinaryNode b)
        {
            spine.Add(b);
            current = b.Left;
        }

        var sb = new System.Text.StringBuilder();
        foreach (var b in spine)
            sb.Append('(').Append(BinaryText(b.Operator)).Append(' ');
        sb.Append(Dump(current, ignoreParentheses));
        for (var i = spine.Count - 1; i >= 0; i--)
            sb.Append(' ').Append(Dump(spine[i].Right, ignoreParentheses)).Append(')');
        return sb.ToString();
    }

    private static string Args(IReadOnlyList<FormulaNode> args, bool ignoreParentheses) =>
        string.Concat(args.Select(a => " " + Dump(a, ignoreParentheses)));

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
