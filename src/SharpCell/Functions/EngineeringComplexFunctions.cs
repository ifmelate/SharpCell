using System;
using System.Globalization;
using System.Text;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// COMPLEX and the IM functions. Excel has no complex type: a complex number is text such as
/// "3+4i" or "1-2j", and every result is written back as text with 15 significant digits.
/// The formulas are the textbook ones Excel uses (IMSQRT and IMPOWER go through polar form,
/// which is why IMSQRT("-4") is "1.22464679914735E-16+2i" in Excel too).
/// </summary>
internal static class EngineeringComplexFunctions
{
    private const int SignificantDigits = 15;

    // A part is written in fixed notation when its 15 significant digits need at most this many
    // decimal places. Excel writes 1E-18 as "0.000000000000000001" but 1E-20 as "1E-20" and
    // 1.22464679914735E-16 in scientific notation; 18 fits all three, the exact limit is not known.
    private const int MaxFixedDecimals = 18;

    private static readonly ArgumentKind[] Values = [ArgumentKind.Value];
    private static readonly ArgumentKind[] AnyArguments = [ArgumentKind.Any];

    public static void Register(FunctionRegistry registry)
    {
        var max = FunctionRegistry.MaxArguments;
        registry.Add(new FunctionInfo("COMPLEX", 2, 3, Values, Complex));
        registry.Add(new FunctionInfo("IMREAL", 1, 1, Values, call => Part(call, z => CellValue.Number(z.Re))));
        registry.Add(new FunctionInfo("IMAGINARY", 1, 1, Values, call => Part(call, z => CellValue.Number(z.Im))));
        registry.Add(new FunctionInfo("IMABS", 1, 1, Values, call => Part(call, z => CellValue.Number(z.Abs))));
        registry.Add(new FunctionInfo("IMARGUMENT", 1, 1, Values, call => Part(call, z =>
            z.IsZero ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(z.Arg))));

        AddUnary(registry, "IMCONJUGATE", z => new Cx(z.Re, -z.Im));
        AddUnary(registry, "IMEXP", z => Exp(z));
        AddUnary(registry, "IMLN", z => Ln(z));
        AddUnary(registry, "IMLOG10", z => Ln(z) / Math.Log(10));
        AddUnary(registry, "IMLOG2", z => Ln(z) / Math.Log(2));
        AddUnary(registry, "IMSQRT", z => Polar(Math.Sqrt(z.Abs), z.Arg / 2));
        AddUnary(registry, "IMSIN", z => Sin(z));
        AddUnary(registry, "IMCOS", z => Cos(z));
        AddUnary(registry, "IMTAN", z => Tan(z, cotangent: false));
        AddUnary(registry, "IMCOT", z => Tan(z, cotangent: true));
        AddUnary(registry, "IMSEC", z => Divide(Cx.One, Cos(z)));
        AddUnary(registry, "IMCSC", z => Divide(Cx.One, Sin(z)));
        AddUnary(registry, "IMSINH", z => Sinh(z));
        AddUnary(registry, "IMCOSH", z => Cosh(z));
        AddUnary(registry, "IMSECH", z => Divide(Cx.One, Cosh(z)));
        AddUnary(registry, "IMCSCH", z => Divide(Cx.One, Sinh(z)));

        registry.Add(new FunctionInfo("IMPOWER", 2, 2, Values, Power));
        registry.Add(new FunctionInfo("IMDIV", 2, 2, Values, call => Binary(call, Divide)));
        registry.Add(new FunctionInfo("IMSUB", 2, 2, Values, call => Binary(call, (a, b) => new Cx(a.Re - b.Re, a.Im - b.Im))));
        registry.Add(new FunctionInfo("IMSUM", 1, max, AnyArguments, call => Fold(call, Cx.Zero, (a, b) => new Cx(a.Re + b.Re, a.Im + b.Im))));
        registry.Add(new FunctionInfo("IMPRODUCT", 1, max, AnyArguments, call => Fold(call, Cx.One, Multiply)));
    }

    private static void AddUnary(FunctionRegistry registry, string name, Func<Cx, Cx?> operation) =>
        registry.Add(new FunctionInfo(name, 1, 1, Values, call =>
        {
            var error = Read(call.Value(0), call.Context.Culture, out var z, out var suffix);
            return error.IsError ? error : Write(operation(z), suffix, call.Context.Culture);
        }));

    private static Operand Complex(FunctionCall call)
    {
        var real = EngineeringFunctions.StrictNumber(call, 0);
        if (real.IsError)
            return real;
        var imaginary = EngineeringFunctions.StrictNumber(call, 1);
        if (imaginary.IsError)
            return imaginary;

        var suffix = 'i';
        if (call.Has(2))
        {
            var value = call.Value(2);
            if (value.IsError)
                return value;
            if (value.Kind == CellValueKind.Text && value.AsText() is "i" or "j")
                suffix = value.AsText()[0];
            else if (!(value.Kind == CellValueKind.Empty || (value.Kind == CellValueKind.Text && value.AsText().Length == 0)))
                return CellValue.Error(ErrorKind.Value);
        }

        return Write(new Cx(real.AsNumber(), imaginary.AsNumber()), suffix, call.Context.Culture);
    }

    private static Operand Part(FunctionCall call, Func<Cx, CellValue> part)
    {
        var error = Read(call.Value(0), call.Context.Culture, out var z, out _);
        return error.IsError ? error : part(z);
    }

    // The exponent must be a number; 0 to a power is 0 for positive powers only.
    private static Operand Power(FunctionCall call)
    {
        var error = Read(call.Value(0), call.Context.Culture, out var z, out var suffix);
        if (error.IsError)
            return error;
        var power = EngineeringFunctions.StrictNumber(call, 1);
        if (power.IsError)
            return power;

        var n = power.AsNumber();
        if (z.IsZero)
            return n > 0 ? Write(Cx.Zero, suffix, call.Context.Culture) : CellValue.Error(ErrorKind.Num);
        return Write(Polar(Math.Pow(z.Abs, n), z.Arg * n), suffix, call.Context.Culture);
    }

    private static Operand Binary(FunctionCall call, Func<Cx, Cx, Cx?> operation)
    {
        var culture = call.Context.Culture;
        var error = Read(call.Value(0), culture, out var a, out var first);
        if (error.IsError)
            return error;
        error = Read(call.Value(1), culture, out var b, out var second);
        if (error.IsError)
            return error;
        return Combine(first, second, out var suffix) ? Write(operation(a, b), suffix, culture) : CellValue.Error(ErrorKind.Value);
    }

    // IMSUM and IMPRODUCT: every value of every argument, ranges and arrays included. An empty
    // cell counts as 0 (so it makes a product 0), like an empty cell given to IMSUB.
    private static Operand Fold(FunctionCall call, Cx seed, Func<Cx, Cx, Cx> operation)
    {
        var culture = call.Context.Culture;
        var result = seed;
        var suffix = '\0';
        CellValue? error = null;
        long referencedCells = 0;
        long visitedCells = 0;
        for (var i = 0; i < call.Count; i++)
        {
            if (call[i].Reference is { } reference)
            {
                foreach (var (_, area) in reference.Areas)
                    referencedCells += area.CellCount;
            }
        }

        Aggregation.ForEach(call, (value, source) =>
        {
            if (source == ValueSource.Reference)
                visitedCells++;
            var failure = Read(value, culture, out var z, out var own);
            if (failure.IsError || !Combine(suffix, own, out suffix))
            {
                error = failure.IsError ? failure : CellValue.Error(ErrorKind.Value);
                return false;
            }

            result = operation(result, z);
            return true;
        });

        if (error is { } e)
            return e;
        if (visitedCells < referencedCells)
            result = operation(result, Cx.Zero);
        return Write(result, suffix, culture);
    }

    // Two suffixes agree when they are equal or one side is a plain real number.
    private static bool Combine(char a, char b, out char suffix)
    {
        suffix = a == '\0' ? b : a;
        return a == '\0' || b == '\0' || a == b;
    }

    /// <summary>
    /// Reads a complex argument. Numbers are real; an empty cell or empty text is 0; logical values
    /// are <c>#VALUE!</c>; text that is not a complex number is <c>#NUM!</c>. The suffix is
    /// <c>'\0'</c> when the text had none.
    /// </summary>
    private static CellValue Read(CellValue value, CultureInfo culture, out Cx z, out char suffix)
    {
        z = Cx.Zero;
        suffix = '\0';
        switch (value.Kind)
        {
            case CellValueKind.Error:
                return value;
            case CellValueKind.Empty:
            case CellValueKind.Missing:
                return CellValue.Empty;
            case CellValueKind.Number:
                z = new Cx(value.AsNumber(), 0);
                return CellValue.Empty;
            case CellValueKind.Text:
                return TryParse(value.AsText(), culture, out z, out suffix) ? CellValue.Empty : CellValue.Error(ErrorKind.Num);
            default:
                return CellValue.Error(ErrorKind.Value);
        }
    }

    // "a", "bi", "a+bi", "a-bi", with "i" or "j", and b left out for 1: "i", "-j", "3+i".
    // Numbers have an optional exponent; there are no spaces anywhere.
    private static bool TryParse(string text, CultureInfo culture, out Cx z, out char suffix)
    {
        z = Cx.Zero;
        suffix = '\0';
        if (text.Length == 0)
            return true;

        var separator = culture.NumberFormat.NumberDecimalSeparator;
        var last = text[^1];
        if (last != 'i' && last != 'j')
        {
            if (!TryParseNumber(text, separator, out var real))
                return false;
            z = new Cx(real, 0);
            return true;
        }

        suffix = last;
        var body = text.AsSpan(0, text.Length - 1);

        // The imaginary part starts at the last sign that is not the sign of an exponent.
        var split = -1;
        for (var i = body.Length - 1; i > 0; i--)
        {
            if (body[i] is '+' or '-' && body[i - 1] is not ('e' or 'E'))
            {
                split = i;
                break;
            }
        }

        var re = 0.0;
        if (split > 0 && !TryParseNumber(body[..split], separator, out re))
            return false;

        var imaginary = split > 0 ? body[split..] : body;
        double im;
        if (imaginary.Length == 0 || imaginary is "+")
            im = 1;
        else if (imaginary is "-")
            im = -1;
        else if (!TryParseNumber(imaginary, separator, out im))
            return false;

        z = new Cx(re, im);
        return true;
    }

    // An optional sign, digits with an optional decimal separator ('.' or the culture's), an
    // optional exponent.
    private static bool TryParseNumber(ReadOnlySpan<char> s, string separator, out double value)
    {
        value = 0;
        var invariant = new StringBuilder(s.Length);
        var i = 0;
        if (i < s.Length && s[i] is '+' or '-')
            invariant.Append(s[i++]);

        var digits = 0;
        while (i < s.Length && char.IsAsciiDigit(s[i]))
        {
            invariant.Append(s[i++]);
            digits++;
        }

        if (i < s.Length && (s[i] == '.' || s[i..].StartsWith(separator, StringComparison.Ordinal)))
        {
            i += s[i] == '.' ? 1 : separator.Length;
            invariant.Append('.');
            while (i < s.Length && char.IsAsciiDigit(s[i]))
            {
                invariant.Append(s[i++]);
                digits++;
            }
        }

        if (digits == 0)
            return false;

        if (i < s.Length && s[i] is 'e' or 'E')
        {
            invariant.Append('e');
            i++;
            if (i < s.Length && s[i] is '+' or '-')
                invariant.Append(s[i++]);
            var exponentDigits = 0;
            while (i < s.Length && char.IsAsciiDigit(s[i]))
            {
                invariant.Append(s[i++]);
                exponentDigits++;
            }

            if (exponentDigits == 0)
                return false;
        }

        return i == s.Length
            && double.TryParse(invariant.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out value)
            && double.IsFinite(value);
    }

    /// <summary>Writes a result as Excel does; no result (a division by zero) and overflow are <c>#NUM!</c>.</summary>
    private static CellValue Write(Cx? result, char suffix, CultureInfo culture)
    {
        if (result is not { } z || !double.IsFinite(z.Re) || !double.IsFinite(z.Im))
            return CellValue.Error(ErrorKind.Num);
        if (z.IsZero)
            return CellValue.Text("0");

        var separator = culture.NumberFormat.NumberDecimalSeparator;
        var sb = new StringBuilder();
        if (z.Re != 0)
            sb.Append(FormatPart(z.Re, separator));
        if (z.Im != 0)
        {
            var magnitude = FormatPart(Math.Abs(z.Im), separator);
            sb.Append(z.Im < 0 ? "-" : z.Re != 0 ? "+" : "");
            if (magnitude != "1")
                sb.Append(magnitude);
            sb.Append(suffix == '\0' ? 'i' : suffix);
        }

        return CellValue.Text(sb.ToString());
    }

    // 15 significant digits, in fixed notation below 1E+15 when that needs at most 18 decimal places.
    private static string FormatPart(double x, string separator)
    {
        var rounded = x.ToString("E" + (SignificantDigits - 1), CultureInfo.InvariantCulture);
        var ePos = rounded.IndexOf('E', StringComparison.Ordinal);
        var exponent = int.Parse(rounded.AsSpan(ePos + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
        var negative = rounded[0] == '-';
        var digits = rounded[(negative ? 1 : 0)..ePos].Replace(".", "", StringComparison.Ordinal).TrimEnd('0');

        var sb = new StringBuilder();
        if (negative)
            sb.Append('-');
        if (exponent >= SignificantDigits || digits.Length - 1 - exponent > MaxFixedDecimals)
        {
            sb.Append(digits[0]);
            if (digits.Length > 1)
                sb.Append(separator).Append(digits, 1, digits.Length - 1);
            sb.Append('E').Append(exponent < 0 ? '-' : '+')
                .Append(Math.Abs(exponent).ToString("00", CultureInfo.InvariantCulture));
        }
        else if (exponent >= 0)
        {
            var integerDigits = exponent + 1;
            if (digits.Length <= integerDigits)
                sb.Append(digits).Append('0', integerDigits - digits.Length);
            else
                sb.Append(digits, 0, integerDigits).Append(separator).Append(digits, integerDigits, digits.Length - integerDigits);
        }
        else
        {
            sb.Append('0').Append(separator).Append('0', -exponent - 1).Append(digits);
        }

        return sb.ToString();
    }

    private static Cx Polar(double r, double theta) => new(r * Math.Cos(theta), r * Math.Sin(theta));

    private static Cx Multiply(Cx a, Cx b) => new(a.Re * b.Re - a.Im * b.Im, a.Re * b.Im + a.Im * b.Re);

    private static Cx? Divide(Cx a, Cx b)
    {
        var d = b.Re * b.Re + b.Im * b.Im;
        if (d == 0)
            return null;
        return new Cx((a.Re * b.Re + a.Im * b.Im) / d, (a.Im * b.Re - a.Re * b.Im) / d);
    }

    private static Cx Exp(Cx z) => Polar(Math.Exp(z.Re), z.Im);

    private static Cx? Ln(Cx z) => z.IsZero ? null : new Cx(Math.Log(z.Abs), z.Arg);

    private static Cx Sin(Cx z) => new(Math.Sin(z.Re) * Math.Cosh(z.Im), Math.Cos(z.Re) * Math.Sinh(z.Im));

    private static Cx Cos(Cx z) => new(Math.Cos(z.Re) * Math.Cosh(z.Im), -Math.Sin(z.Re) * Math.Sinh(z.Im));

    // tan z = (sin 2x + i sinh 2y) / (cos 2x + cosh 2y) and cot z = (sin 2x - i sinh 2y) / (cosh 2y - cos 2x):
    // dividing sin z by cos z instead loses the real part to cancellation once y is large. The
    // denominators are evaluated as 2 (sinh² y + sin² x) and 2 (sinh² y + cos² x), which keeps
    // cot 1E-16 at 1E+16.
    private static Cx? Tan(Cx z, bool cotangent)
    {
        var sinhY = Math.Sinh(z.Im);
        var trig = cotangent ? Math.Sin(z.Re) : Math.Cos(z.Re);
        var d = 2 * (sinhY * sinhY + trig * trig);
        if (d == 0)
            return null;
        var im = Math.Sinh(2 * z.Im) / d;
        return new Cx(Math.Sin(2 * z.Re) / d, cotangent ? -im : im);
    }

    private static Cx Sinh(Cx z) => new(Math.Sinh(z.Re) * Math.Cos(z.Im), Math.Cosh(z.Re) * Math.Sin(z.Im));

    private static Cx Cosh(Cx z) => new(Math.Cosh(z.Re) * Math.Cos(z.Im), Math.Sinh(z.Re) * Math.Sin(z.Im));

    private readonly record struct Cx(double Re, double Im)
    {
        public static Cx Zero => new(0, 0);

        public static Cx One => new(1, 0);

        public bool IsZero => Re == 0 && Im == 0;

        public double Abs => Math.Sqrt(Re * Re + Im * Im);

        public double Arg => Math.Atan2(Im, Re);

        public static Cx? operator /(Cx? z, double d) => z is { } v ? new Cx(v.Re / d, v.Im / d) : null;
    }
}
