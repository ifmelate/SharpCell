using System;
using System.Collections.Generic;
using System.Text;

namespace SharpCell.Parsing;

/// <summary>
/// Prints a formula tree as canonical text (en-US syntax, no leading <c>=</c>). Explicit
/// <see cref="ParenthesesNode"/>s are printed as written; any other parentheses are added only
/// where the tree would otherwise parse differently.
/// </summary>
internal sealed class FormulaPrinter
{
    // Binding strength, loosest first. Binary operators are left-associative: the left operand
    // needs at least the operator's precedence, the right operand strictly more.
    private const int Comparison = 1;
    private const int Concat = 2;
    private const int Additive = 3;
    private const int Multiplicative = 4;
    private const int Power = 5;
    private const int Percent = 6;
    private const int Prefix = 7;
    private const int Union = 8;
    private const int Intersect = 9;
    private const int ImplicitIntersection = 10;
    private const int Range = 11;
    private const int Primary = 12;

    private readonly StringBuilder _sb = new();
    private readonly CellAddress _origin;
    private readonly ReferenceStyle _style;

    private FormulaPrinter(CellAddress origin, ReferenceStyle style)
    {
        _origin = origin;
        _style = style;
    }

    /// <param name="origin">The cell the formula belongs to; relative references are printed from it.</param>
    public static string Print(FormulaNode node, CellAddress origin, ReferenceStyle style = ReferenceStyle.A1)
    {
        ArgumentNullException.ThrowIfNull(node);
        var printer = new FormulaPrinter(origin, style);
        printer.Write(node, 0, unionAllowed: false);
        return printer._sb.ToString();
    }

    private void Write(FormulaNode node, int minPrecedence, bool unionAllowed)
    {
        if (NeedsParentheses(node, minPrecedence, unionAllowed))
        {
            _sb.Append('(');
            WriteBare(node, unionAllowed: true);
            _sb.Append(')');
            return;
        }

        WriteBare(node, unionAllowed);
    }

    private static bool NeedsParentheses(FormulaNode node, int minPrecedence, bool unionAllowed) =>
        PrecedenceOf(node) < minPrecedence || (node is BinaryNode { Operator: BinaryOperator.Union } && !unionAllowed);

    private void WriteBare(FormulaNode node, bool unionAllowed)
    {
        switch (node)
        {
            case NumberNode n:
                _sb.Append(NumberText.FormatLiteral(n.Value));
                break;
            case TextNode t:
                WriteString(t.Value);
                break;
            case BooleanNode b:
                _sb.Append(b.Value ? "TRUE" : "FALSE");
                break;
            case ErrorNode e:
                _sb.Append(e.Error.ToText());
                break;
            case MissingNode:
                break;
            case ReferenceNode r:
                WriteSheet(r.Sheet);
                _sb.Append(_style == ReferenceStyle.R1C1
                    ? ReferenceSyntax.FormatR1C1(r.Area)
                    : ReferenceSyntax.FormatA1(r.Area, _origin));
                break;
            case RefErrorNode r:
                WriteSheet(r.Sheet);
                _sb.Append(ErrorKind.Ref.ToText());
                break;
            case NameNode n:
                WriteSheet(n.Sheet);
                _sb.Append(n.Name);
                break;
            case StructuredReferenceNode s:
                _sb.Append(s.Text);
                break;
            case UnaryNode { Operator: UnaryOperator.Percent } u:
                Write(u.Operand, Percent, unionAllowed);
                _sb.Append('%');
                break;
            case UnaryNode u:
                _sb.Append(u.Operator == UnaryOperator.Negate ? '-' : '+');
                Write(u.Operand, Prefix, unionAllowed);
                break;
            case BinaryNode b:
                WriteBinary(b, unionAllowed);
                break;
            case FunctionNode f:
                _sb.Append(f.Name);
                WriteArguments(f.Arguments);
                break;
            case CallNode c:
                if (c.Callee is FunctionNode or CallNode or ParenthesesNode)
                    WriteBare(c.Callee, unionAllowed);
                else
                    Write(c.Callee, Primary + 1, unionAllowed);
                WriteArguments(c.Arguments);
                break;
            case ParenthesesNode p:
                _sb.Append('(');
                Write(p.Inner, 0, unionAllowed: true);
                _sb.Append(')');
                break;
            case ArrayNode a:
                WriteArray(a.Values);
                break;
            case SpillNode s:
                Write(s.Operand, Primary, unionAllowed);
                _sb.Append('#');
                break;
            case ImplicitIntersectionNode i:
                _sb.Append('@');
                Write(i.Operand, ImplicitIntersection, unionAllowed);
                break;
            default:
                throw new ArgumentException($"Unknown node type {node.GetType().Name}.", nameof(node));
        }
    }

    // Parsed chains like A1+A1+...+A1 are left-deep trees thousands of nodes tall. The left spine
    // is walked in a loop so printing does not recurse once per operator.
    private void WriteBinary(BinaryNode node, bool unionAllowed)
    {
        var chain = new List<BinaryNode> { node };
        var current = node;
        while (current.Left is BinaryNode left && !NeedsParentheses(left, PrecedenceOf(current), unionAllowed))
        {
            chain.Add(left);
            current = left;
        }

        Write(current.Left, PrecedenceOf(current), unionAllowed);
        for (var i = chain.Count - 1; i >= 0; i--)
        {
            var b = chain[i];
            var operatorAt = _sb.Length;
            _sb.Append(OperatorText(b.Operator));

            // A space means intersection only when a reference-like operand follows it.
            if (b.Operator == BinaryOperator.Intersect && b.Right is NumberNode or TextNode or BooleanNode or ErrorNode or ArrayNode)
            {
                _sb.Append('(');
                WriteBare(b.Right, unionAllowed: true);
                _sb.Append(')');
            }
            else
            {
                Write(b.Right, PrecedenceOf(b) + 1, unionAllowed);
            }

            if (b.Operator == BinaryOperator.Range && !FusesToSameArea(b) && !LexesAsColon(operatorAt))
                _sb.Insert(operatorAt, ' ');
        }
    }

    // Two cells fuse into the same area the parser builds from "A1 : B2", so no space is needed.
    private static bool FusesToSameArea(BinaryNode range) =>
        range.Left is ReferenceNode { Area.Kind: AreaKind.Cell } && range.Right is ReferenceNode { Sheet: null, Area.Kind: AreaKind.Cell };

    // The lexer reads words across a colon ("1:1" rows, "A:B" columns, "X:Y!" sheets), so text on
    // both sides of a range operator can fuse into one token: Range(1, rows 1:3) would print as
    // "1:1:3". The printed text is checked with the lexer itself; a space before ':' prevents fusion.
    private bool LexesAsColon(int position)
    {
        try
        {
            foreach (var token in Lexer.Tokenize(_sb.ToString(), _origin, _style))
            {
                if (token.Start == position)
                    return token.Kind == TokenKind.Colon;
                if (token.Start > position)
                    return false;
            }

            return false;
        }
        catch (FormulaParseException)
        {
            return false;
        }
    }

    private void WriteArguments(IReadOnlyList<FormulaNode> arguments)
    {
        _sb.Append('(');
        for (var i = 0; i < arguments.Count; i++)
        {
            if (i > 0)
                _sb.Append(',');
            Write(arguments[i], 0, unionAllowed: false);
        }

        _sb.Append(')');
    }

    private void WriteArray(CellValue[,] values)
    {
        _sb.Append('{');
        for (var r = 0; r < values.GetLength(0); r++)
        {
            if (r > 0)
                _sb.Append(';');
            for (var c = 0; c < values.GetLength(1); c++)
            {
                if (c > 0)
                    _sb.Append(',');
                var value = values[r, c];
                switch (value.Kind)
                {
                    case CellValueKind.Number:
                        _sb.Append(NumberText.FormatLiteral(value.AsNumber()));
                        break;
                    case CellValueKind.Text:
                        WriteString(value.AsText());
                        break;
                    default:
                        _sb.Append(value.ToString());
                        break;
                }
            }
        }

        _sb.Append('}');
    }

    private void WriteString(string value) =>
        _sb.Append('"').Append(value.Replace("\"", "\"\"", StringComparison.Ordinal)).Append('"');

    private void WriteSheet(SheetPrefix? sheet)
    {
        if (sheet is null)
            return;

        var quote = NeedsQuotes(sheet.First) || (sheet.Last is not null && NeedsQuotes(sheet.Last));
        var text = sheet.Last is null ? sheet.First : sheet.First + ":" + sheet.Last;
        if (quote)
            _sb.Append('\'').Append(text.Replace("'", "''", StringComparison.Ordinal)).Append('\'');
        else
            _sb.Append(text);
        _sb.Append('!');
    }

    // Unquoted sheet names must lex as one word and must not read as a reference or a boolean.
    private static bool NeedsQuotes(string name)
    {
        if (name.Length == 0 || !(char.IsLetter(name[0]) || name[0] == '_'))
            return true;

        foreach (var c in name)
        {
            if (!(char.IsLetterOrDigit(c) || c is '_' or '.'))
                return true;
        }

        return CellAddress.TryParse(name, out _)
            || ReferenceSyntax.TryParseR1C1Area(name, out _)
            || name.Equals("TRUE", StringComparison.OrdinalIgnoreCase)
            || name.Equals("FALSE", StringComparison.OrdinalIgnoreCase);
    }

    private static int PrecedenceOf(FormulaNode node) => node switch
    {
        BinaryNode b => b.Operator switch
        {
            BinaryOperator.Range => Range,
            BinaryOperator.Intersect => Intersect,
            BinaryOperator.Union => Union,
            BinaryOperator.Power => Power,
            BinaryOperator.Multiply or BinaryOperator.Divide => Multiplicative,
            BinaryOperator.Add or BinaryOperator.Subtract => Additive,
            BinaryOperator.Concat => Concat,
            _ => Comparison,
        },
        UnaryNode { Operator: UnaryOperator.Percent } => Percent,
        UnaryNode => Prefix,

        // A negative literal prints with a leading '-' and reads back as a prefix operator.
        NumberNode { Value: < 0 } => Prefix,
        ImplicitIntersectionNode => ImplicitIntersection,
        _ => Primary,
    };

    private static string OperatorText(BinaryOperator op) => op switch
    {
        BinaryOperator.Range => ":",
        BinaryOperator.Intersect => " ",
        BinaryOperator.Union => ",",
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
