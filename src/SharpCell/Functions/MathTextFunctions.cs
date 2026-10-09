using System;
using System.Text;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>Numbers to and from text: ROMAN, ARABIC, BASE and DECIMAL.</summary>
internal static class MathTextFunctions
{
    private static readonly ArgumentKind[] ValueArguments = [ArgumentKind.Value];

    private const string Digits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    // BASE and DECIMAL work on integers below 2^53 and texts of at most 255 characters.
    private const double MaxInteger = 9007199254740992;
    private const int MaxLength = 255;

    private static readonly char[] RomanLetters = ['M', 'D', 'C', 'L', 'X', 'V', 'I'];
    private static readonly int[] RomanValues = [1000, 500, 100, 50, 10, 5, 1];

    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("ROMAN", 1, 2, ValueArguments, Roman));
        registry.Add(new FunctionInfo("ARABIC", 1, 1, ValueArguments, Arabic));
        registry.Add(new FunctionInfo("BASE", 2, 3, ValueArguments, Base));
        registry.Add(new FunctionInfo("DECIMAL", 2, 2, ValueArguments, Decimal));
    }

    // Forms 0 (classic, the default or TRUE) to 4 (simplest, or FALSE) allow ever more
    // subtractive pairs: 499 is CDXCIX, LDVLIV, XDIX, VDIV, ID.
    private static Operand Roman(FunctionCall call)
    {
        var number = call.Number(0);
        if (number.IsError)
            return number;
        var mode = 0.0;
        if (call.Has(1))
        {
            var form = call.Value(1);
            if (form.Kind == CellValueKind.Boolean)
            {
                mode = form.AsBoolean() ? 0 : 4;
            }
            else
            {
                var converted = call.Integer(1);
                if (converted.IsError)
                    return converted;
                mode = converted.AsNumber();
            }
        }

        var value = Math.Truncate(number.AsNumber());
        if (value < 0 || value >= 4000 || mode < 0 || mode > 4)
            return CellValue.Error(ErrorKind.Value);
        return CellValue.Text(ToRoman((int)value, (int)mode));
    }

    // Each decimal digit is written from the letters of its decade; a 4 or 9 becomes a pair whose
    // smaller letter is pushed up to `mode` letters further down (the scheme Excel and
    // LibreOffice share).
    private static string ToRoman(int value, int mode)
    {
        var text = new StringBuilder();
        for (var decade = 0; decade < RomanValues.Length; decade += 2)
        {
            var digit = value / RomanValues[decade];
            if (digit % 5 == 4)
            {
                var larger = digit == 4 ? decade - 1 : decade - 2;
                var smaller = decade;
                for (var steps = 0; steps < mode && smaller < RomanValues.Length - 1; steps++)
                {
                    if (RomanValues[larger] - RomanValues[smaller + 1] > value)
                        break;
                    smaller++;
                }

                text.Append(RomanLetters[smaller]).Append(RomanLetters[larger]);
                value += RomanValues[smaller] - RomanValues[larger];
            }
            else
            {
                if (digit > 4)
                    text.Append(RomanLetters[decade - 1]);
                text.Append(RomanLetters[decade], digit % 5);
                value %= RomanValues[decade];
            }
        }

        return text.ToString();
    }

    // Any sequence of numerals reads: a letter smaller than the next one is subtracted, so IM is
    // 999 and XIXI 20. Case and surrounding spaces do not matter; a leading minus negates.
    private static Operand Arabic(FunctionCall call)
    {
        var value = call.Value(0);
        if (value.IsError)
            return value;
        if (value.Kind is CellValueKind.Empty or CellValueKind.Missing)
            return CellValue.Number(0);
        if (value.Kind != CellValueKind.Text)
            return CellValue.Error(ErrorKind.Value);

        var text = value.AsText().Trim();
        if (text.Length > MaxLength)
            return CellValue.Error(ErrorKind.Value);
        var negative = text.StartsWith('-');
        if (negative)
            text = text[1..];

        var total = 0;
        for (var i = 0; i < text.Length; i++)
        {
            var current = RomanValue(text[i]);
            if (current == 0)
                return CellValue.Error(ErrorKind.Value);
            var next = i + 1 < text.Length ? RomanValue(text[i + 1]) : 0;
            total += next > current ? -current : current;
        }

        return CellValue.Number(negative ? -total : total);
    }

    private static int RomanValue(char letter) => char.ToUpperInvariant(letter) switch
    {
        'I' => 1,
        'V' => 5,
        'X' => 10,
        'L' => 50,
        'C' => 100,
        'D' => 500,
        'M' => 1000,
        _ => 0,
    };

    private static Operand Base(FunctionCall call)
    {
        var number = call.Number(0);
        if (number.IsError)
            return number;
        var radix = call.Integer(1);
        if (radix.IsError)
            return radix;
        var minimum = call.Integer(2, 0);
        if (minimum.IsError)
            return minimum;

        var value = Math.Truncate(number.AsNumber());
        var @base = radix.AsNumber();
        var length = minimum.AsNumber();
        if (value < 0 || value >= MaxInteger || @base < 2 || @base > 36 || length < 0 || length > MaxLength)
            return CellValue.Error(ErrorKind.Num);

        var digits = new StringBuilder();
        var rest = (long)value;
        do
        {
            digits.Insert(0, Digits[(int)(rest % (long)@base)]);
            rest /= (long)@base;
        }
        while (rest > 0);

        if (digits.Length < length)
            digits.Insert(0, "0", (int)length - digits.Length);
        return CellValue.Text(digits.ToString());
    }

    private static Operand Decimal(FunctionCall call)
    {
        var text = call.Text(0);
        if (text.IsError)
            return text;
        var radix = call.Integer(1);
        if (radix.IsError)
            return radix;

        var digits = text.AsText();
        var @base = radix.AsNumber();
        if (digits.Length > MaxLength || @base < 2 || @base > 36)
            return CellValue.Error(ErrorKind.Num);

        var total = 0.0;
        foreach (var c in digits)
        {
            var digit = Digits.IndexOf(char.ToUpperInvariant(c), StringComparison.Ordinal);
            if (digit < 0 || digit >= @base)
                return CellValue.Error(ErrorKind.Num);
            total = total * @base + digit;
        }

        return total >= MaxInteger ? CellValue.Error(ErrorKind.Num) : CellValue.Number(total);
    }
}
