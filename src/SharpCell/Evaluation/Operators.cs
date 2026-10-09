using System;
using System.Collections.Generic;
using System.Globalization;
using SharpCell.Parsing;

namespace SharpCell.Evaluation;

/// <summary>Excel semantics of the formula operators on scalar values and on references.</summary>
internal static class Operators
{
    /// <summary>Excel's limit on text length; longer results are <c>#VALUE!</c>.</summary>
    public const int MaxTextLength = 32_767;

    public static CellValue Negate(CellValue value, CultureInfo culture)
    {
        var number = Coercion.ToNumber(value, culture);
        return number.IsError ? number : CellValue.Number(-number.AsNumber());
    }

    public static CellValue Percent(CellValue value, CultureInfo culture)
    {
        var number = Coercion.ToNumber(value, culture);
        return number.IsError ? number : CellValue.Number(number.AsNumber() / 100);
    }

    /// <param name="last">
    /// Whether this is the formula's last operation. Like Excel, a last addition or subtraction whose
    /// operands cancel out up to floating-point noise gives exactly 0 (=0.3-0.2-0.1 is 0).
    /// </param>
    public static CellValue Binary(BinaryOperator op, CellValue left, CellValue right, CultureInfo culture, bool last)
    {
        switch (op)
        {
            case BinaryOperator.Concat:
                var leftText = Coercion.ToText(left, culture);
                if (leftText.IsError)
                    return leftText;
                var rightText = Coercion.ToText(right, culture);
                if (rightText.IsError)
                    return rightText;
                if ((long)leftText.AsText().Length + rightText.AsText().Length > MaxTextLength)
                    return CellValue.Error(ErrorKind.Value);
                return CellValue.Text(leftText.AsText() + rightText.AsText());
            case BinaryOperator.Equal:
            case BinaryOperator.NotEqual:
            case BinaryOperator.Less:
            case BinaryOperator.LessEqual:
            case BinaryOperator.Greater:
            case BinaryOperator.GreaterEqual:
                return Compare(op, left, right, culture);
        }

        var x = Coercion.ToNumber(left, culture);
        if (x.IsError)
            return x;
        var y = Coercion.ToNumber(right, culture);
        if (y.IsError)
            return y;

        double a = x.AsNumber(), b = y.AsNumber();
        switch (op)
        {
            case BinaryOperator.Add:
                return CellValue.Number(last && NumberComparer.AreEqual(a, -b) ? 0 : a + b);
            case BinaryOperator.Subtract:
                return CellValue.Number(last && NumberComparer.AreEqual(a, b) ? 0 : a - b);
            case BinaryOperator.Multiply:
                return CellValue.Number(a * b);
            case BinaryOperator.Divide:
                return b == 0 ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(a / b);
            case BinaryOperator.Power:
                if (a == 0 && b == 0)
                    return CellValue.Error(ErrorKind.Num);
                if (a == 0 && b < 0)
                    return CellValue.Error(ErrorKind.Div0);
                if (a < 0 && b != Math.Floor(b))
                    return OddRoot(a, b);
                return CellValue.Number(Math.Pow(a, b));
            default:
                throw new ArgumentOutOfRangeException(nameof(op), op, "Not a value operator.");
        }
    }

    // A negative base with a fractional exponent has a real result only for odd roots: Excel
    // gives (-8)^(1/3) = -2 and #NUM! for (-8)^(1/2) or (-8)^0.4.
    private static CellValue OddRoot(double a, double b)
    {
        var degree = Math.Round(1 / b);
        if (Math.Abs(1 / b - degree) > 1e-9 * Math.Abs(degree) || Math.Abs(degree % 2) != 1)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(-Math.Pow(-a, b));
    }

    // Values are never coerced for comparison. Types order as number < text < boolean, and an
    // empty cell stands for 0, "" or FALSE, whichever the other side is.
    private static CellValue Compare(BinaryOperator op, CellValue left, CellValue right, CultureInfo culture)
    {
        if (left.IsError)
            return left;
        if (right.IsError)
            return right;

        left = FillEmpty(left, right);
        right = FillEmpty(right, left);
        int order;
        var leftRank = Rank(left);
        var rightRank = Rank(right);
        if (leftRank != rightRank)
        {
            order = leftRank.CompareTo(rightRank);
        }
        else
        {
            order = left.Kind switch
            {
                CellValueKind.Number => NumberComparer.Compare(left.AsNumber(), right.AsNumber()),
                CellValueKind.Text => TextComparer.Compare(left.AsText(), right.AsText(), culture),
                CellValueKind.Boolean => left.AsBoolean().CompareTo(right.AsBoolean()),
                _ => 0,
            };
        }

        return CellValue.Boolean(op switch
        {
            BinaryOperator.Equal => order == 0,
            BinaryOperator.NotEqual => order != 0,
            BinaryOperator.Less => order < 0,
            BinaryOperator.LessEqual => order <= 0,
            BinaryOperator.Greater => order > 0,
            _ => order >= 0,
        });
    }

    private static CellValue FillEmpty(CellValue value, CellValue other)
    {
        if (value.Kind is not (CellValueKind.Empty or CellValueKind.Missing))
            return value;

        return other.Kind switch
        {
            CellValueKind.Text => CellValue.Text(""),
            CellValueKind.Boolean => CellValue.False,
            CellValueKind.Empty or CellValueKind.Missing => value,
            _ => CellValue.Number(0),
        };
    }

    private static int Rank(CellValue value) => value.Kind switch
    {
        CellValueKind.Number => 0,
        CellValueKind.Text => 1,
        CellValueKind.Boolean => 2,
        _ => 3,
    };

    /// <summary><c>A1:B2</c> on references: the bounding box of both sides, which must be on one sheet.</summary>
    public static Operand Range(Operand left, Operand right)
    {
        if (left.Reference is not { } a || right.Reference is not { } b)
            return CellValue.Error(ErrorKind.Value);

        var sheet = a.Areas[0].Sheet;
        var box = a.Areas[0].Area;
        foreach (var part in (IEnumerable<SheetArea>)[.. a.Areas, .. b.Areas])
        {
            if (part.Sheet != sheet)
                return CellValue.Error(ErrorKind.Value);
            box = Area.Bounding(box, part.Area);
        }

        return Operand.Of(new Reference(sheet, box));
    }

    /// <summary>The space operator: the cells both references share, or <c>#NULL!</c>.</summary>
    public static Operand Intersect(Operand left, Operand right)
    {
        if (left.Reference is not { } a || right.Reference is not { } b)
            return CellValue.Error(ErrorKind.Value);

        var result = new List<SheetArea>();
        foreach (var x in a.Areas)
        {
            foreach (var y in b.Areas)
            {
                if (x.Sheet == y.Sheet && x.Area.TryIntersect(y.Area, out var common))
                    result.Add(new SheetArea(x.Sheet, common));
            }
        }

        return result.Count == 0 ? CellValue.Error(ErrorKind.Null) : Operand.Of(new Reference(result));
    }

    /// <summary>The comma operator inside parentheses: all areas of both references.</summary>
    public static Operand Union(Operand left, Operand right)
    {
        if (left.Reference is not { } a || right.Reference is not { } b)
            return CellValue.Error(ErrorKind.Value);

        return Operand.Of(new Reference([.. a.Areas, .. b.Areas]));
    }
}
