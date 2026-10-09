using System;
using System.Text;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>DELTA, GESTEP, the BIT functions and the conversions between number bases.</summary>
internal static class EngineeringFunctions
{
    // BIT functions work on integers below 2^48.
    private const double BitLimit = 281474976710656;

    // The largest shift Excel accepts either way.
    private const int MaxShift = 53;

    // Base conversions write at most ten digits; a negative number is the ten-digit two's complement.
    private const int MaxDigits = 10;

    private static readonly ArgumentKind[] Values = [ArgumentKind.Value];

    private static readonly NumberBase Binary = new(2);
    private static readonly NumberBase Octal = new(8);
    private static readonly NumberBase Hexadecimal = new(16);

    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("DELTA", 1, 2, Values, call => Compare(call, (a, b) => a == b)));
        registry.Add(new FunctionInfo("GESTEP", 1, 2, Values, call => Compare(call, (a, b) => a >= b)));

        registry.Add(new FunctionInfo("BITAND", 2, 2, Values, call => Bitwise(call, (a, b) => a & b)));
        registry.Add(new FunctionInfo("BITOR", 2, 2, Values, call => Bitwise(call, (a, b) => a | b)));
        registry.Add(new FunctionInfo("BITXOR", 2, 2, Values, call => Bitwise(call, (a, b) => a ^ b)));
        registry.Add(new FunctionInfo("BITLSHIFT", 2, 2, Values, call => Shift(call, left: true)));
        registry.Add(new FunctionInfo("BITRSHIFT", 2, 2, Values, call => Shift(call, left: false)));

        registry.Add(new FunctionInfo("DEC2BIN", 1, 2, Values, call => FromDecimal(call, Binary)));
        registry.Add(new FunctionInfo("DEC2OCT", 1, 2, Values, call => FromDecimal(call, Octal)));
        registry.Add(new FunctionInfo("DEC2HEX", 1, 2, Values, call => FromDecimal(call, Hexadecimal)));
        AddBase(registry, "BIN", Binary);
        AddBase(registry, "OCT", Octal);
        AddBase(registry, "HEX", Hexadecimal);
    }

    // BIN2DEC, BIN2OCT, BIN2HEX and the same for the other two bases.
    private static void AddBase(FunctionRegistry registry, string prefix, NumberBase from)
    {
        registry.Add(new FunctionInfo(prefix + "2DEC", 1, 1, Values, call => ToDecimal(call, from)));
        foreach (var (suffix, to) in new[] { ("BIN", Binary), ("OCT", Octal), ("HEX", Hexadecimal) })
        {
            if (to != from)
                registry.Add(new FunctionInfo(prefix + "2" + suffix, 1, 2, Values, call => Between(call, from, to)));
        }
    }

    /// <summary>
    /// A number argument the way the Analysis ToolPak functions read it: numbers and numeric text
    /// convert, an empty cell is 0, but a logical value is <c>#VALUE!</c>, unlike in arithmetic.
    /// </summary>
    internal static CellValue StrictNumber(FunctionCall call, int index)
    {
        var value = call.Value(index);
        return value.Kind == CellValueKind.Boolean
            ? CellValue.Error(ErrorKind.Value)
            : Coercion.ToNumber(value, call.Context.Culture, call.Context.DateSystem);
    }

    internal static CellValue StrictNumber(FunctionCall call, int index, double absent) =>
        call.Has(index) ? StrictNumber(call, index) : CellValue.Number(absent);

    private static Operand Compare(FunctionCall call, Func<double, double, bool> test)
    {
        var a = StrictNumber(call, 0);
        if (a.IsError)
            return a;
        var b = StrictNumber(call, 1, 0);
        if (b.IsError)
            return b;
        return CellValue.Number(test(a.AsNumber(), b.AsNumber()) ? 1 : 0);
    }

    // An operand of the BIT functions: a whole number in [0, 2^48). Logical values count as 1 and 0.
    private static CellValue BitOperand(FunctionCall call, int index)
    {
        var number = call.Number(index);
        if (number.IsError)
            return number;
        var n = number.AsNumber();
        return n < 0 || n >= BitLimit || n != Math.Floor(n) ? CellValue.Error(ErrorKind.Num) : number;
    }

    private static Operand Bitwise(FunctionCall call, Func<long, long, long> operation)
    {
        var a = BitOperand(call, 0);
        if (a.IsError)
            return a;
        var b = BitOperand(call, 1);
        if (b.IsError)
            return b;
        return CellValue.Number(operation((long)a.AsNumber(), (long)b.AsNumber()));
    }

    // A negative shift goes the other way. The result must still fit in 48 bits.
    private static Operand Shift(FunctionCall call, bool left)
    {
        var number = BitOperand(call, 0);
        if (number.IsError)
            return number;
        var amount = call.Integer(1);
        if (amount.IsError)
            return amount;
        var shift = amount.AsNumber();
        if (Math.Abs(shift) > MaxShift)
            return CellValue.Error(ErrorKind.Num);

        var bits = (long)number.AsNumber();
        var toLeft = left ? (int)shift : -(int)shift;
        if (toLeft <= 0)
            return CellValue.Number(bits >> -toLeft);
        var result = bits * Math.Pow(2, toLeft);
        return result >= BitLimit ? CellValue.Error(ErrorKind.Num) : CellValue.Number(result);
    }

    private static Operand FromDecimal(FunctionCall call, NumberBase to)
    {
        var number = StrictNumber(call, 0);
        if (number.IsError)
            return number;
        var places = Places(call);
        if (places.IsError)
            return places;
        return Format(Math.Truncate(number.AsNumber()), to, places);
    }

    private static Operand ToDecimal(FunctionCall call, NumberBase from)
    {
        var digits = DigitsArgument(call);
        if (digits.IsError)
            return digits;
        return Parse(digits.AsText(), from);
    }

    private static Operand Between(FunctionCall call, NumberBase from, NumberBase to)
    {
        var digits = DigitsArgument(call);
        if (digits.IsError)
            return digits;
        var number = Parse(digits.AsText(), from);
        if (number.IsError)
            return number;
        var places = Places(call);
        if (places.IsError)
            return places;
        return Format(number.AsNumber(), to, places);
    }

    // The number to convert is read as text: 1100100 typed as a number means the digits "1100100".
    private static CellValue DigitsArgument(FunctionCall call)
    {
        var value = call.Value(0);
        return value.Kind == CellValueKind.Boolean ? CellValue.Error(ErrorKind.Value) : Coercion.ToText(value, call.Context.Culture);
    }

    // Places, when given, must be 1 to 10 (fractions truncated), whether or not the result needs it.
    private static CellValue Places(FunctionCall call)
    {
        if (!call.Has(1))
            return CellValue.Empty;
        var places = StrictNumber(call, 1);
        if (places.IsError)
            return places;
        var n = Math.Truncate(places.AsNumber());
        return n is < 1 or > MaxDigits ? CellValue.Error(ErrorKind.Num) : CellValue.Number(n);
    }

    // At most ten digits; with all ten, a leading digit in the upper half of the base makes the
    // number negative (two's complement). An empty text is 0.
    private static CellValue Parse(string text, NumberBase from)
    {
        if (text.Length > MaxDigits)
            return CellValue.Error(ErrorKind.Num);

        long value = 0;
        foreach (var c in text)
        {
            var digit = DigitValue(c);
            if (digit < 0 || digit >= from.Radix)
                return CellValue.Error(ErrorKind.Num);
            value = value * from.Radix + digit;
        }

        if (text.Length == MaxDigits && DigitValue(text[0]) >= from.Radix / 2)
            value -= from.Modulus;
        return CellValue.Number(value);
    }

    private static int DigitValue(char c) => c switch
    {
        >= '0' and <= '9' => c - '0',
        >= 'A' and <= 'F' => c - 'A' + 10,
        >= 'a' and <= 'f' => c - 'a' + 10,
        _ => -1,
    };

    // A positive number is padded to places (and must fit); a negative one ignores places.
    private static CellValue Format(double number, NumberBase to, CellValue places)
    {
        if (number < -to.Modulus / 2 || number >= to.Modulus / 2)
            return CellValue.Error(ErrorKind.Num);

        var value = (long)number;
        if (value < 0)
            value += to.Modulus;

        var sb = new StringBuilder();
        do
        {
            sb.Insert(0, "0123456789ABCDEF"[(int)(value % to.Radix)]);
            value /= to.Radix;
        }
        while (value > 0);

        if (number >= 0 && places.Kind == CellValueKind.Number)
        {
            var width = (int)places.AsNumber();
            if (sb.Length > width)
                return CellValue.Error(ErrorKind.Num);
            sb.Insert(0, "0", width - sb.Length);
        }

        return CellValue.Text(sb.ToString());
    }

    private sealed record NumberBase(int Radix)
    {
        // Radix^10: the ten-digit range, half of it negative.
        public long Modulus { get; } = (long)Math.Pow(Radix, MaxDigits);
    }
}
