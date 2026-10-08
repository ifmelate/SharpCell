using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace SharpCell.Parsing;

/// <summary>
/// Recursive-descent parser for canonical (en-US) formula text. Precedence, from loosest:
/// comparison, <c>&amp;</c>, <c>+ -</c>, <c>* /</c>, <c>^</c>, <c>%</c>, unary <c>- +</c>,
/// union <c>,</c>, intersection (space), <c>@</c>, range <c>:</c>, postfix <c>#</c> and calls.
/// All binary operators are left-associative and built in loops; only nesting (parentheses,
/// calls, arrays, prefix operators) recurses, capped by <see cref="FormulaLimits.MaxDepth"/>.
/// </summary>
internal sealed class FormulaParser
{
    private readonly List<Token> _tokens;
    private int _index;
    private int _depth;

    // A comma is the union operator only inside parentheses; elsewhere it separates arguments.
    private bool _unionAllowed;

    private FormulaParser(List<Token> tokens)
    {
        _tokens = tokens;
    }

    /// <summary>Parses formula text with an optional leading <c>=</c>.</summary>
    /// <param name="origin">The cell the formula belongs to; relative references are stored as offsets from it.</param>
    public static FormulaNode Parse(string text, CellAddress origin, ReferenceStyle style = ReferenceStyle.A1)
    {
        ArgumentNullException.ThrowIfNull(text);
        var start = text.StartsWith('=') ? 1 : 0;
        var parser = new FormulaParser(Lexer.Tokenize(text, origin, style, start));
        var node = parser.ParseExpression();
        if (parser.Current.Kind != TokenKind.End)
            throw Unexpected(parser.Current);
        return node;
    }

    private Token Current => _tokens[_index];

    private Token Advance() => _tokens[_index++];

    private FormulaNode ParseExpression() => ParseBinary(1);

    // Precedence climbing over the arithmetic, text and comparison operators. A chain of
    // same-level operators is built in the loop, so only a higher-precedence right operand recurses.
    private FormulaNode ParseBinary(int minPrecedence)
    {
        var left = ParsePrefixed();
        while (true)
        {
            var (op, precedence) = BinaryOperatorOf(Current.Kind);
            if (precedence < minPrecedence)
                return left;

            Advance();
            left = new BinaryNode(op, left, ParseBinary(precedence + 1));
        }
    }

    private static (BinaryOperator Operator, int Precedence) BinaryOperatorOf(TokenKind kind) => kind switch
    {
        TokenKind.Equal => (BinaryOperator.Equal, 1),
        TokenKind.NotEqual => (BinaryOperator.NotEqual, 1),
        TokenKind.Less => (BinaryOperator.Less, 1),
        TokenKind.LessEqual => (BinaryOperator.LessEqual, 1),
        TokenKind.Greater => (BinaryOperator.Greater, 1),
        TokenKind.GreaterEqual => (BinaryOperator.GreaterEqual, 1),
        TokenKind.Ampersand => (BinaryOperator.Concat, 2),
        TokenKind.Plus => (BinaryOperator.Add, 3),
        TokenKind.Minus => (BinaryOperator.Subtract, 3),
        TokenKind.Star => (BinaryOperator.Multiply, 4),
        TokenKind.Slash => (BinaryOperator.Divide, 4),
        TokenKind.Caret => (BinaryOperator.Power, 5),
        _ => (default, 0),
    };

    // Unary signs bind looser than reference operators but tighter than '%' and '^':
    // -2^2 is (-2)^2 and -5% is (-5)%.
    private FormulaNode ParsePrefixed()
    {
        var signs = new List<UnaryOperator>();
        while (Current.Kind is TokenKind.Minus or TokenKind.Plus)
        {
            Enter();
            signs.Add(Advance().Kind == TokenKind.Minus ? UnaryOperator.Negate : UnaryOperator.Plus);
        }

        var node = ParseReferenceExpression(1);
        for (var i = signs.Count - 1; i >= 0; i--)
        {
            node = new UnaryNode(signs[i], node);
            Leave();
        }

        while (Current.Kind == TokenKind.Percent)
        {
            Advance();
            node = new UnaryNode(UnaryOperator.Percent, node);
        }

        return node;
    }

    // Union (precedence 1, only inside parentheses) and intersection (2, a space between operands).
    private FormulaNode ParseReferenceExpression(int minPrecedence)
    {
        var left = ParseRangeOperand();
        while (true)
        {
            BinaryOperator op;
            int precedence;
            if (_unionAllowed && Current.Kind == TokenKind.Comma)
                (op, precedence) = (BinaryOperator.Union, 1);
            else if (Current.SpaceBefore && StartsReferenceOperand(Current.Kind))
                (op, precedence) = (BinaryOperator.Intersect, 2);
            else
                return left;

            if (precedence < minPrecedence)
                return left;

            if (op == BinaryOperator.Union)
                Advance();
            left = new BinaryNode(op, left, ParseReferenceExpression(precedence + 1));
        }
    }

    // Only operands that can evaluate to a reference make a space mean intersection;
    // "A1 -B1" stays a subtraction and "1 2" stays an error.
    private static bool StartsReferenceOperand(TokenKind kind) => kind is TokenKind.Reference or TokenKind.Name
        or TokenKind.Function or TokenKind.StructuredReference or TokenKind.OpenParen or TokenKind.At;

    // '@' applies to a whole range expression: @A1:A10.
    private FormulaNode ParseRangeOperand()
    {
        var ats = 0;
        while (Current.Kind == TokenKind.At)
        {
            Enter();
            Advance();
            ats++;
        }

        var node = ParsePostfix();
        while (Current.Kind == TokenKind.Colon)
        {
            Advance();
            node = new BinaryNode(BinaryOperator.Range, node, ParsePostfix());
        }

        for (var i = 0; i < ats; i++)
        {
            node = new ImplicitIntersectionNode(node);
            Leave();
        }

        return node;
    }

    private FormulaNode ParsePostfix()
    {
        var node = ParsePrimary();
        while (true)
        {
            if (Current.Kind == TokenKind.Hash)
            {
                Advance();
                node = new SpillNode(node);
            }
            else if (Current.Kind == TokenKind.OpenParen && !Current.SpaceBefore
                     && node is FunctionNode or CallNode or ParenthesesNode)
            {
                node = new CallNode(node, ParseArguments());
            }
            else
            {
                return node;
            }
        }
    }

    private FormulaNode ParsePrimary()
    {
        var token = Current;
        switch (token.Kind)
        {
            case TokenKind.Number:
                Advance();
                return new NumberNode(token.Number);
            case TokenKind.String:
                Advance();
                return new TextNode(token.Text!);
            case TokenKind.Boolean:
                Advance();
                return new BooleanNode(token.Boolean);
            case TokenKind.Error:
                Advance();
                return token.Sheet is null ? new ErrorNode(token.Error) : new RefErrorNode(token.Sheet);
            case TokenKind.Reference:
                Advance();
                return new ReferenceNode(token.Sheet, token.Area);
            case TokenKind.Name:
                Advance();
                return new NameNode(token.Sheet, StripParameterPrefix(token.Text!));
            case TokenKind.StructuredReference:
                Advance();
                return new StructuredReferenceNode(token.Text!);
            case TokenKind.Function:
                return ParseFunction();
            case TokenKind.OpenParen:
                return ParseParentheses();
            case TokenKind.OpenBrace:
                return ParseArray();
            default:
                throw Unexpected(token);
        }
    }

    private FormulaNode ParseFunction()
    {
        var token = Advance();
        var name = NormalizeFunctionName(token.Text!);
        var arguments = ParseArguments();
        switch (name)
        {
            case "ANCHORARRAY":
                return new SpillNode(SingleArgument(arguments, token));
            case "SINGLE":
                return new ImplicitIntersectionNode(SingleArgument(arguments, token));
            default:
                return new FunctionNode(name, arguments);
        }
    }

    private static FormulaNode SingleArgument(List<FormulaNode> arguments, Token function)
    {
        if (arguments.Count != 1 || arguments[0] is MissingNode)
            throw new FormulaParseException($"{function.Text} takes exactly one argument.", function.Start);
        return arguments[0];
    }

    private List<FormulaNode> ParseArguments()
    {
        Expect(TokenKind.OpenParen);
        Enter();
        var unionAllowed = _unionAllowed;
        _unionAllowed = false;

        var arguments = new List<FormulaNode>();
        if (Current.Kind == TokenKind.CloseParen)
        {
            Advance();
        }
        else
        {
            while (true)
            {
                arguments.Add(Current.Kind is TokenKind.Comma or TokenKind.CloseParen ? MissingNode.Instance : ParseExpression());
                if (Current.Kind == TokenKind.Comma)
                {
                    Advance();
                    continue;
                }

                Expect(TokenKind.CloseParen);
                break;
            }
        }

        _unionAllowed = unionAllowed;
        Leave();
        return arguments;
    }

    private FormulaNode ParseParentheses()
    {
        Advance();
        Enter();
        var unionAllowed = _unionAllowed;
        _unionAllowed = true;
        var inner = ParseExpression();
        _unionAllowed = unionAllowed;
        Expect(TokenKind.CloseParen);
        Leave();
        return new ParenthesesNode(inner);
    }

    private FormulaNode ParseArray()
    {
        Advance();
        Enter();
        var rows = new List<List<CellValue>>();
        var row = new List<CellValue>();
        while (true)
        {
            row.Add(ParseArrayElement());
            var separator = Current;
            if (separator.Kind == TokenKind.Comma)
            {
                Advance();
                continue;
            }

            if (separator.Kind is not (TokenKind.Semicolon or TokenKind.CloseBrace))
                throw Unexpected(separator);

            if (rows.Count > 0 && row.Count != rows[0].Count)
                throw new FormulaParseException("Array rows must have the same number of columns.", separator.Start);

            rows.Add(row);
            Advance();
            if (separator.Kind == TokenKind.CloseBrace)
                break;
            row = [];
        }

        Leave();
        var values = new CellValue[rows.Count, rows[0].Count];
        for (var r = 0; r < rows.Count; r++)
        {
            for (var c = 0; c < rows[r].Count; c++)
                values[r, c] = rows[r][c];
        }

        return new ArrayNode(values);
    }

    private CellValue ParseArrayElement()
    {
        var token = Current;
        if (token.Kind is TokenKind.Minus or TokenKind.Plus)
        {
            Advance();
            var number = Current;
            if (number.Kind != TokenKind.Number)
                throw Unexpected(number);
            Advance();
            return CellValue.Number(token.Kind == TokenKind.Minus ? -number.Number : number.Number);
        }

        switch (token.Kind)
        {
            case TokenKind.Number:
                Advance();
                return CellValue.Number(token.Number);
            case TokenKind.String:
                Advance();
                return CellValue.Text(token.Text!);
            case TokenKind.Boolean:
                Advance();
                return CellValue.Boolean(token.Boolean);
            case TokenKind.Error when token.Sheet is null:
                Advance();
                return CellValue.Error(token.Error);
            default:
                throw Unexpected(token);
        }
    }

    private void Expect(TokenKind kind)
    {
        if (Current.Kind != kind)
            throw Unexpected(Current);
        Advance();
    }

    // The depth limit covers normal threads; the stack probe covers callers on small stacks.
    private void Enter()
    {
        if (++_depth > FormulaLimits.MaxDepth || !RuntimeHelpers.TryEnsureSufficientExecutionStack())
            throw new FormulaParseException("Formula is nested too deeply.", Current.Start);
    }

    private void Leave() => _depth--;

    private static string NormalizeFunctionName(string name)
    {
        var upper = name.ToUpperInvariant();
        while (true)
        {
            if (upper.StartsWith("_XLFN.", StringComparison.Ordinal))
                upper = upper[6..];
            else if (upper.StartsWith("_XLWS.", StringComparison.Ordinal))
                upper = upper[6..];
            else
                return upper;
        }
    }

    // Files store LAMBDA/LET parameter names as "_xlpm.x".
    private static string StripParameterPrefix(string name) =>
        name.StartsWith("_xlpm.", StringComparison.OrdinalIgnoreCase) ? name[6..] : name;

    private static FormulaParseException Unexpected(Token token) => token.Kind == TokenKind.End
        ? new FormulaParseException("Unexpected end of formula.", token.Start)
        : new FormulaParseException("Unexpected token.", token.Start);
}
