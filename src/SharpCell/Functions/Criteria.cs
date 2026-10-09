using System;
using System.Globalization;

namespace SharpCell.Functions;

/// <summary>
/// A condition of COUNTIF, SUMIF, the *IFS functions and the database functions, as Excel reads
/// it: <c>10</c>, <c>"&gt;=10"</c>, <c>"&lt;&gt;apple"</c>, <c>"a*"</c>, <c>"~*"</c>, <c>TRUE</c>, <c>#N/A</c>.
/// <para>
/// The criterion's operand decides the kind of comparison. A number compares with numbers; with
/// <c>=</c> it also matches text that reads as that number, while <c>&lt;&gt;</c> counts all text as
/// different. TRUE/FALSE (also as text, any case) compare with booleans, error values with
/// errors. Other text compares with text: <c>=</c> and <c>&lt;&gt;</c> with wildcards (<c>*</c>, <c>?</c>,
/// <c>~</c> escapes) ignoring case, the ordering operators by the culture's order. <c>&lt;&gt;</c>
/// matches every cell the same criterion with <c>=</c> does not, empty cells included. An empty
/// criterion cell means 0.
/// </para>
/// </summary>
internal sealed class Criterion
{
    private enum Operator
    {
        Equal,
        NotEqual,
        Less,
        LessEqual,
        Greater,
        GreaterEqual,
    }

    private enum OperandKind
    {
        Number,
        Boolean,
        Error,
        Text,
    }

    private readonly Operator _op;
    private readonly OperandKind _kind;
    private readonly double _number;
    private readonly bool _boolean;
    private readonly ErrorKind _error;
    private readonly string _text = "";
    private readonly bool _hasWildcards;
    private readonly CultureInfo _culture;
    private readonly DateSystem _dateSystem;

    private Criterion(Operator op, OperandKind kind, CultureInfo culture, DateSystem dateSystem)
    {
        _op = op;
        _kind = kind;
        _culture = culture;
        _dateSystem = dateSystem;
    }

    private Criterion(Operator op, double number, CultureInfo culture, DateSystem dateSystem)
        : this(op, OperandKind.Number, culture, dateSystem) => _number = number;

    private Criterion(Operator op, bool boolean, CultureInfo culture, DateSystem dateSystem)
        : this(op, OperandKind.Boolean, culture, dateSystem) => _boolean = boolean;

    private Criterion(Operator op, ErrorKind error, CultureInfo culture, DateSystem dateSystem)
        : this(op, OperandKind.Error, culture, dateSystem) => _error = error;

    private Criterion(Operator op, string text, CultureInfo culture, DateSystem dateSystem)
        : this(op, OperandKind.Text, culture, dateSystem)
    {
        _text = text;
        _hasWildcards = text.AsSpan().IndexOfAny("*?~") >= 0;
    }

    /// <summary>Whether an empty cell meets the criterion; ranges count their empty cells without visiting them.</summary>
    public bool MatchesEmpty => Matches(CellValue.Empty);

    public static Criterion Parse(CellValue criterion, CultureInfo culture, DateSystem dateSystem)
    {
        switch (criterion.Kind)
        {
            case CellValueKind.Number:
                return new Criterion(Operator.Equal, criterion.AsNumber(), culture, dateSystem);
            case CellValueKind.Boolean:
                return new Criterion(Operator.Equal, criterion.AsBoolean(), culture, dateSystem);
            case CellValueKind.Error:
                return new Criterion(Operator.Equal, criterion.AsError(), culture, dateSystem);
            case CellValueKind.Text:
                return ParseText(criterion.AsText(), culture, dateSystem);
            default:
                return new Criterion(Operator.Equal, 0.0, culture, dateSystem);
        }
    }

    private static Criterion ParseText(string text, CultureInfo culture, DateSystem dateSystem)
    {
        var (op, length) = text switch
        {
            _ when text.StartsWith("<=", StringComparison.Ordinal) => (Operator.LessEqual, 2),
            _ when text.StartsWith(">=", StringComparison.Ordinal) => (Operator.GreaterEqual, 2),
            _ when text.StartsWith("<>", StringComparison.Ordinal) => (Operator.NotEqual, 2),
            _ when text.StartsWith('<') => (Operator.Less, 1),
            _ when text.StartsWith('>') => (Operator.Greater, 1),
            _ when text.StartsWith('=') => (Operator.Equal, 1),
            _ => (Operator.Equal, 0),
        };
        var operand = text[length..];

        if (operand.Length > 0)
        {
            var number = Coercion.ToNumber(CellValue.Text(operand), culture, dateSystem);
            if (number.Kind == CellValueKind.Number)
                return new Criterion(op, number.AsNumber(), culture, dateSystem);
            if (operand.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
                return new Criterion(op, true, culture, dateSystem);
            if (operand.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
                return new Criterion(op, false, culture, dateSystem);
            if (ErrorKinds.TryParse(operand, out var error))
                return new Criterion(op, error, culture, dateSystem);
        }

        return new Criterion(op, operand, culture, dateSystem);
    }

    public bool Matches(CellValue value)
    {
        if (_op == Operator.NotEqual)
            return !Equal(value);
        if (_op == Operator.Equal)
            return Equal(value);

        // Ordering compares like with like only.
        int order;
        switch (_kind)
        {
            case OperandKind.Number when value.Kind == CellValueKind.Number:
                order = value.AsNumber().CompareTo(_number);
                break;
            // "<=", ">" and the like with nothing after the operator match nothing, not even "".
            case OperandKind.Text when _text.Length == 0:
                return false;
            case OperandKind.Text when value.Kind == CellValueKind.Text:
                order = string.Compare(value.AsText(), _text, _culture, CompareOptions.IgnoreCase);
                break;
            case OperandKind.Boolean when value.Kind == CellValueKind.Boolean:
                order = value.AsBoolean().CompareTo(_boolean);
                break;
            default:
                return false;
        }

        return _op switch
        {
            Operator.Less => order < 0,
            Operator.LessEqual => order <= 0,
            Operator.Greater => order > 0,
            _ => order >= 0,
        };
    }

    private bool Equal(CellValue value)
    {
        switch (_kind)
        {
            case OperandKind.Number:
                if (value.Kind == CellValueKind.Number)
                    return value.AsNumber() == _number;

                // "=10" also finds the text "10", but "<>10" counts it as different (Excel's asymmetry).
                return _op == Operator.Equal && value.Kind == CellValueKind.Text
                    && NumberText.TryParse(value.AsText(), _culture, out var parsed) && parsed == _number;
            case OperandKind.Boolean:
                return value.Kind == CellValueKind.Boolean && value.AsBoolean() == _boolean;
            case OperandKind.Error:
                return value.Kind == CellValueKind.Error && value.AsError() == _error;
            default:
                if (_text.Length == 0)
                    return _op == Operator.NotEqual
                        ? value.Kind == CellValueKind.Empty
                        : value.Kind == CellValueKind.Empty || (value.Kind == CellValueKind.Text && value.AsText().Length == 0);
                if (value.Kind != CellValueKind.Text)
                    return false;
                return _hasWildcards
                    ? Wildcard.Matches(_text, value.AsText(), _culture)
                    : string.Compare(value.AsText(), _text, _culture, CompareOptions.IgnoreCase) == 0;
        }
    }
}

/// <summary>Excel wildcards: <c>*</c> any run, <c>?</c> one character, <c>~</c> escapes the next one. Case is ignored.</summary>
internal static class Wildcard
{
    public static bool Matches(string pattern, string text, CultureInfo culture)
    {
        var p = culture.TextInfo.ToUpper(pattern);
        var t = culture.TextInfo.ToUpper(text);

        // Greedy matching with one backtrack point per '*': linear in practice, no regex.
        int pi = 0, ti = 0, starPattern = -1, starText = 0;
        while (ti < t.Length)
        {
            if (pi < p.Length && p[pi] == '*')
            {
                starPattern = ++pi;
                starText = ti;
                continue;
            }

            if (pi < p.Length && Literal(p, ref pi, out var expected, out var any) && (any || expected == t[ti]))
            {
                ti++;
                continue;
            }

            if (starPattern < 0)
                return false;
            pi = starPattern;
            ti = ++starText;
        }

        while (pi < p.Length && p[pi] == '*')
            pi++;
        return pi == p.Length;
    }

    // Reads one pattern position: '?' (any character), '~x' (literal x) or a plain character.
    // Leaves pi unchanged unless the position was read.
    private static bool Literal(string p, ref int pi, out char expected, out bool any)
    {
        expected = '\0';
        any = false;
        if (p[pi] == '?')
        {
            any = true;
            pi++;
            return true;
        }

        if (p[pi] == '~' && pi + 1 < p.Length && p[pi + 1] is '*' or '?' or '~')
        {
            expected = p[pi + 1];
            pi += 2;
            return true;
        }

        expected = p[pi];
        pi++;
        return true;
    }
}
