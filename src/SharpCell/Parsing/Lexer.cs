using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SharpCell.Parsing;

/// <summary>
/// Splits formula text (canonical en-US syntax, without the leading <c>=</c>) into tokens.
/// References are resolved to <see cref="AreaRef"/> here because telling <c>A1</c> from a name or
/// <c>1:3</c> from a number needs the same character-level look-ahead.
/// </summary>
internal sealed class Lexer
{
    private readonly string _text;
    private readonly CellAddress _origin;
    private readonly ReferenceStyle _style;
    private readonly List<Token> _tokens = [];
    private int _pos;
    private bool _spaceBefore;

    private Lexer(string text, CellAddress origin, ReferenceStyle style)
    {
        _text = text;
        _origin = origin;
        _style = style;
    }

    /// <param name="start">Index to start at, e.g. 1 to skip a leading <c>=</c>; positions stay relative to <paramref name="text"/>.</param>
    public static List<Token> Tokenize(string text, CellAddress origin, ReferenceStyle style, int start = 0)
    {
        if (text.Length > FormulaLimits.MaxLength)
            throw new FormulaParseException($"Formula is longer than {FormulaLimits.MaxLength} characters.", FormulaLimits.MaxLength);

        var lexer = new Lexer(text, origin, style) { _pos = start };
        lexer.Run();
        return lexer._tokens;
    }

    private void Run()
    {
        while (true)
        {
            _spaceBefore = false;
            while (_pos < _text.Length && IsWhitespace(_text[_pos]))
            {
                _spaceBefore = true;
                _pos++;
            }

            if (_pos >= _text.Length)
            {
                Add(new Token(TokenKind.End, _pos, 0));
                return;
            }

            LexToken();
        }
    }

    private void LexToken()
    {
        var start = _pos;
        var c = _text[start];
        switch (c)
        {
            case '"':
                LexString();
                return;
            case '#':
                LexHash();
                return;
            case '\'':
                LexQuotedPrefix();
                return;
            case '[':
                LexBracket();
                return;
            case '<':
                if (Peek(1) == '>')
                    Punct(TokenKind.NotEqual, 2);
                else if (Peek(1) == '=')
                    Punct(TokenKind.LessEqual, 2);
                else
                    Punct(TokenKind.Less, 1);
                return;
            case '>':
                Punct(Peek(1) == '=' ? TokenKind.GreaterEqual : TokenKind.Greater, Peek(1) == '=' ? 2 : 1);
                return;
        }

        var punct = c switch
        {
            '+' => TokenKind.Plus,
            '-' => TokenKind.Minus,
            '*' => TokenKind.Star,
            '/' => TokenKind.Slash,
            '^' => TokenKind.Caret,
            '&' => TokenKind.Ampersand,
            '%' => TokenKind.Percent,
            '=' => TokenKind.Equal,
            ':' => TokenKind.Colon,
            ',' => TokenKind.Comma,
            ';' => TokenKind.Semicolon,
            '(' => TokenKind.OpenParen,
            ')' => TokenKind.CloseParen,
            '{' => TokenKind.OpenBrace,
            '}' => TokenKind.CloseBrace,
            '@' => TokenKind.At,
            _ => TokenKind.End,
        };
        if (punct != TokenKind.End)
        {
            Punct(punct, 1);
            return;
        }

        if (char.IsAsciiDigit(c) || c == '.')
        {
            if (_style == ReferenceStyle.A1 && TryLexReference(start, null))
                return;
            LexNumber();
            return;
        }

        if (IsNameStart(c))
        {
            LexWord();
            return;
        }

        throw Error($"Unexpected character '{c}'.", start);
    }

    private void LexString()
    {
        var start = _pos;
        var sb = new StringBuilder();
        var i = start + 1;
        while (true)
        {
            if (i >= _text.Length)
                throw Error("Unterminated string.", start);

            if (_text[i] == '"')
            {
                if (i + 1 < _text.Length && _text[i + 1] == '"')
                {
                    sb.Append('"');
                    i += 2;
                    continue;
                }

                break;
            }

            sb.Append(_text[i]);
            i++;
        }

        _pos = i + 1;
        Add(new Token(TokenKind.String, start, _pos - start) { Text = sb.ToString() });
    }

    // '#' starts an error literal; otherwise it is the spill operator (A1#).
    private void LexHash()
    {
        var start = _pos;
        if (TryMatchError(start, out var kind, out var length))
        {
            _pos += length;
            Add(new Token(TokenKind.Error, start, length) { Error = kind });
            return;
        }

        Punct(TokenKind.Hash, 1);
    }

    private void LexQuotedPrefix()
    {
        var start = _pos;
        var sb = new StringBuilder();
        var i = start + 1;
        while (true)
        {
            if (i >= _text.Length)
                throw Error("Unterminated quoted sheet name.", start);

            if (_text[i] == '\'')
            {
                if (i + 1 < _text.Length && _text[i + 1] == '\'')
                {
                    sb.Append('\'');
                    i += 2;
                    continue;
                }

                break;
            }

            sb.Append(_text[i]);
            i++;
        }

        i++;
        if (i >= _text.Length || _text[i] != '!')
            throw Error("Expected '!' after quoted sheet name.", start);

        var content = sb.ToString();
        if (content.StartsWith('['))
            throw Error("External workbook references are not supported.", start);

        // Sheet names cannot contain ':', so a colon always separates the two ends of a 3D reference.
        var colon = content.IndexOf(':', StringComparison.Ordinal);
        var prefix = colon < 0
            ? new SheetPrefix(content)
            : new SheetPrefix(content[..colon], content[(colon + 1)..]);
        if (prefix.First.Length == 0 || prefix.Last is { Length: 0 })
            throw Error("Empty sheet name.", start);

        LexAfterPrefix(start, i + 1, prefix);
    }

    private void LexBracket()
    {
        var start = _pos;
        var i = start + 1;
        while (i < _text.Length && char.IsAsciiDigit(_text[i]))
            i++;
        if (i > start + 1 && i < _text.Length && _text[i] == ']')
            throw Error("External workbook references are not supported.", start);

        LexStructured(start, start);
    }

    private void LexStructured(int start, int bracket)
    {
        var depth = 0;
        var i = bracket;
        while (i < _text.Length)
        {
            var c = _text[i];
            if (c == '\'')
            {
                // An apostrophe escapes the next character inside a structured reference.
                i += 2;
                continue;
            }

            if (c == '[')
            {
                depth++;
            }
            else if (c == ']')
            {
                depth--;
                if (depth == 0)
                {
                    _pos = i + 1;
                    Add(new Token(TokenKind.StructuredReference, start, _pos - start) { Text = _text[start.._pos] });
                    return;
                }
            }

            i++;
        }

        throw Error("Unterminated structured reference.", start);
    }

    private void LexNumber()
    {
        var start = _pos;
        var i = start;
        var digits = 0;
        while (i < _text.Length && char.IsAsciiDigit(_text[i]))
        {
            i++;
            digits++;
        }

        if (i < _text.Length && _text[i] == '.')
        {
            i++;
            while (i < _text.Length && char.IsAsciiDigit(_text[i]))
            {
                i++;
                digits++;
            }
        }

        if (digits == 0)
            throw Error("Malformed number.", start);

        if (i < _text.Length && (_text[i] is 'e' or 'E'))
        {
            i++;
            if (i < _text.Length && (_text[i] is '+' or '-'))
                i++;
            var exponentStart = i;
            while (i < _text.Length && char.IsAsciiDigit(_text[i]))
                i++;
            if (i == exponentStart)
                throw Error("Malformed number exponent.", start);
        }

        if (i < _text.Length && IsNameChar(_text[i]))
            throw Error("Malformed number.", start);

        if (!double.TryParse(_text.AsSpan(start, i - start), NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent,
                CultureInfo.InvariantCulture, out var value) || !double.IsFinite(value))
            throw Error("Number is out of range.", start);

        _pos = i;
        Add(new Token(TokenKind.Number, start, i - start) { Number = value });
    }

    private void LexWord()
    {
        var start = _pos;
        if (_style == ReferenceStyle.R1C1 && TryLexR1C1Reference(start, null))
            return;

        var end = ScanWord(start);
        var word = _text[start..end];
        var next = Peek(end - start);

        if (next == '!')
        {
            LexAfterPrefix(start, end + 1, new SheetPrefix(ValidateUnquotedSheet(word, start)));
            return;
        }

        if (next == ':')
        {
            var end2 = ScanWord(end + 1);
            if (end2 > end + 1 && end2 < _text.Length && _text[end2] == '!')
            {
                var prefix = new SheetPrefix(ValidateUnquotedSheet(word, start), ValidateUnquotedSheet(_text[(end + 1)..end2], start));
                LexAfterPrefix(start, end2 + 1, prefix);
                return;
            }
        }

        if (next == '(')
        {
            _pos = end;
            Add(new Token(TokenKind.Function, start, end - start) { Text = ValidateName(word, start) });
            return;
        }

        if (next == '[')
        {
            LexStructured(start, end);
            return;
        }

        if (_style == ReferenceStyle.A1 && TryLexReference(start, null))
            return;

        _pos = end;
        if (word.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || word.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
        {
            Add(new Token(TokenKind.Boolean, start, end - start) { Boolean = word.Length == 4 });
            return;
        }

        Add(new Token(TokenKind.Name, start, end - start) { Text = ValidateName(word, start) });
    }

    // After "Sheet!": a reference, a name, or #REF!. Errors point at the start of the prefix.
    private void LexAfterPrefix(int tokenStart, int pos, SheetPrefix prefix)
    {
        _pos = pos;
        if (pos < _text.Length && _text[pos] == '#')
        {
            if (TryMatchError(pos, out var kind, out var length) && kind == ErrorKind.Ref)
            {
                _pos = pos + length;
                Add(new Token(TokenKind.Error, tokenStart, _pos - tokenStart) { Error = kind, Sheet = prefix });
                return;
            }

            throw Error("Expected a reference after the sheet name.", tokenStart);
        }

        if (_style == ReferenceStyle.R1C1 ? TryLexR1C1Reference(pos, prefix, tokenStart) : TryLexReference(pos, prefix, tokenStart))
            return;

        if (pos < _text.Length && IsNameStart(_text[pos]))
        {
            var end = ScanWord(pos);
            if (end >= _text.Length || (_text[end] is not ('(' or '!' or '[')))
            {
                _pos = end;
                Add(new Token(TokenKind.Name, tokenStart, end - tokenStart) { Text = ValidateName(_text[pos..end], tokenStart), Sheet = prefix });
                return;
            }
        }

        throw Error("Expected a reference after the sheet name.", tokenStart);
    }

    private bool TryLexReference(int pos, SheetPrefix? prefix, int tokenStart = -1)
    {
        if (tokenStart < 0)
            tokenStart = pos;

        var end1 = ScanReferenceWord(pos);
        if (end1 == pos)
            return false;

        if (end1 < _text.Length && _text[end1] == ':')
        {
            var end2 = ScanReferenceWord(end1 + 1);
            if (end2 > end1 + 1 && FollowsReference(end2)
                && ReferenceSyntax.TryParseA1Area(_text.AsSpan(pos, end2 - pos), _origin, out var range))
            {
                AddReference(tokenStart, end2, range, prefix);
                return true;
            }
        }

        if (FollowsReference(end1) && ReferenceSyntax.TryParseA1Area(_text.AsSpan(pos, end1 - pos), _origin, out var cell))
        {
            AddReference(tokenStart, end1, cell, prefix);
            return true;
        }

        return false;
    }

    private bool TryLexR1C1Reference(int pos, SheetPrefix? prefix, int tokenStart = -1)
    {
        if (tokenStart < 0)
            tokenStart = pos;

        var end1 = ScanR1C1Part(pos);
        if (end1 == pos)
            return false;

        if (end1 < _text.Length && _text[end1] == ':')
        {
            var end2 = ScanR1C1Part(end1 + 1);
            if (end2 > end1 + 1 && FollowsReference(end2)
                && ReferenceSyntax.TryParseR1C1Area(_text.AsSpan(pos, end2 - pos), out var range))
            {
                AddReference(tokenStart, end2, range, prefix);
                return true;
            }
        }

        if (FollowsReference(end1) && ReferenceSyntax.TryParseR1C1Area(_text.AsSpan(pos, end1 - pos), out var single))
        {
            AddReference(tokenStart, end1, single, prefix);
            return true;
        }

        return false;
    }

    private void AddReference(int start, int end, AreaRef area, SheetPrefix? prefix)
    {
        _pos = end;
        Add(new Token(TokenKind.Reference, start, end - start) { Area = area, Sheet = prefix });
    }

    // A reference must not run into a name, a call or another prefix: "A1B", "LOG10(", "A1!".
    private bool FollowsReference(int i) =>
        i >= _text.Length || !(IsNameChar(_text[i]) || _text[i] is '(' or '!' or '[');

    private int ScanReferenceWord(int i)
    {
        while (i < _text.Length && (char.IsAsciiLetterOrDigit(_text[i]) || _text[i] == '$'))
            i++;
        return i;
    }

    private int ScanR1C1Part(int i)
    {
        if (i < _text.Length && (_text[i] is 'R' or 'r'))
            i = ScanR1C1Axis(i + 1);
        if (i < _text.Length && (_text[i] is 'C' or 'c'))
            i = ScanR1C1Axis(i + 1);
        return i;
    }

    // Digits, or "[+-digits]". An unclosed bracket is left unconsumed so the reference is rejected.
    private int ScanR1C1Axis(int i)
    {
        if (i < _text.Length && _text[i] == '[')
        {
            var j = i + 1;
            if (j < _text.Length && (_text[j] is '+' or '-'))
                j++;
            while (j < _text.Length && char.IsAsciiDigit(_text[j]))
                j++;
            return j < _text.Length && _text[j] == ']' ? j + 1 : i;
        }

        while (i < _text.Length && char.IsAsciiDigit(_text[i]))
            i++;
        return i;
    }

    private int ScanWord(int i)
    {
        while (i < _text.Length && IsNameChar(_text[i]))
            i++;
        return i;
    }

    private bool TryMatchError(int pos, out ErrorKind kind, out int length) =>
        ErrorKinds.TryMatchPrefix(_text.AsSpan(pos), out kind, out length);

    private static string ValidateUnquotedSheet(string name, int position) =>
        name.Contains('$', StringComparison.Ordinal) ? throw Error("Invalid sheet name.", position) : name;

    private static string ValidateName(string name, int position) =>
        name.Contains('$', StringComparison.Ordinal) ? throw Error($"Invalid name '{name}'.", position) : name;

    private void Punct(TokenKind kind, int length)
    {
        Add(new Token(kind, _pos, length));
        _pos += length;
    }

    private void Add(Token token) => _tokens.Add(token with { SpaceBefore = _spaceBefore });

    private char Peek(int offset) => _pos + offset < _text.Length ? _text[_pos + offset] : '\0';

    private static bool IsWhitespace(char c) => c is ' ' or '\t' or '\r' or '\n';

    private static bool IsNameStart(char c) => char.IsLetter(c) || c is '_' or '\\' or '$';

    private static bool IsNameChar(char c) => char.IsLetterOrDigit(c) || c is '_' or '.' or '\\' or '?' or '$';

    private static FormulaParseException Error(string message, int position) => new(message, position);
}
