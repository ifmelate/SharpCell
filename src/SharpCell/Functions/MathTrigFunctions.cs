using System;

namespace SharpCell.Functions;

/// <summary>Trigonometric and hyperbolic functions, their inverses, DEGREES and RADIANS.</summary>
internal static class MathTrigFunctions
{
    private static readonly ArgumentKind[] ValueArguments = [ArgumentKind.Value];

    // Excel refuses the circular functions of angles from 2^27 up: too few bits are left for the fraction.
    private const double MaxAngle = 134217728;

    // Excel reduces angles by the 64-bit-mantissa pi of the x87 unit, PI() plus this much, not by
    // the exact pi: that is why its SIN(PI()) is 1.22514845490862E-16 rather than 1.2246...E-16.
    private const double PiLow = 1.22514845490862e-16;

    public static void Register(FunctionRegistry registry)
    {
        Add(registry, "SIN", x => Circular(x, Sin));
        Add(registry, "COS", x => Circular(x, Cos));
        Add(registry, "TAN", x => Circular(x, Tan));
        Add(registry, "COT", x => x == 0 ? CellValue.Error(ErrorKind.Div0) : Circular(x, a => 1 / Tan(a)));
        Add(registry, "CSC", x => x == 0 ? CellValue.Error(ErrorKind.Div0) : Circular(x, a => 1 / Sin(a)));
        Add(registry, "SEC", x => Circular(x, a => 1 / Cos(a)));
        Add(registry, "ASIN", x => Math.Abs(x) > 1 ? CellValue.Error(ErrorKind.Num) : Math.Asin(x));
        Add(registry, "ACOS", x => Math.Abs(x) > 1 ? CellValue.Error(ErrorKind.Num) : Math.Acos(x));
        Add(registry, "ATAN", x => Math.Atan(x));

        // Excel's ACOT takes values in (0, pi), so ACOT(-1) is 3pi/4.
        Add(registry, "ACOT", x => x == 0 ? Math.PI / 2 : x > 0 ? Math.Atan(1 / x) : Math.PI + Math.Atan(1 / x));
        Add(registry, "SINH", x => Math.Sinh(x));
        Add(registry, "COSH", x => Math.Cosh(x));
        Add(registry, "TANH", x => Math.Tanh(x));
        Add(registry, "COTH", x => x == 0 ? CellValue.Error(ErrorKind.Div0) : 1 / Math.Tanh(x));
        Add(registry, "CSCH", x => x == 0 ? CellValue.Error(ErrorKind.Div0) : 1 / Math.Sinh(x));
        Add(registry, "SECH", x => 1 / Math.Cosh(x));
        Add(registry, "ASINH", Asinh);
        Add(registry, "ACOSH", x => x < 1 ? CellValue.Error(ErrorKind.Num) : Math.Acosh(x));
        Add(registry, "ATANH", x => Math.Abs(x) >= 1 ? CellValue.Error(ErrorKind.Num) : Math.Atanh(x));
        Add(registry, "ACOTH", x => Math.Abs(x) <= 1 ? CellValue.Error(ErrorKind.Num) : Math.Atanh(1 / x));
        Add(registry, "DEGREES", x => x * (180 / Math.PI));
        Add(registry, "RADIANS", x => x * (Math.PI / 180));
        registry.Add(new FunctionInfo("ATAN2", 2, 2, ValueArguments, MathFunctions.Binary(Atan2)));
    }

    private static void Add(FunctionRegistry registry, string name, Func<double, CellValue> f) =>
        registry.Add(new FunctionInfo(name, 1, 1, ValueArguments, MathFunctions.Unary(f)));

    private static CellValue Circular(double x, Func<double, double> f) =>
        Math.Abs(x) >= MaxAngle ? CellValue.Error(ErrorKind.Num) : CellValue.Number(f(x));

    private static double Sin(double x)
    {
        var (r, quadrant) = Reduce(x);
        return quadrant switch
        {
            0 => Math.Sin(r),
            1 => Math.Cos(r),
            2 => -Math.Sin(r),
            _ => -Math.Cos(r),
        };
    }

    private static double Cos(double x)
    {
        var (r, quadrant) = Reduce(x);
        return quadrant switch
        {
            0 => Math.Cos(r),
            1 => -Math.Sin(r),
            2 => -Math.Cos(r),
            _ => Math.Sin(r),
        };
    }

    private static double Tan(double x)
    {
        var (r, quadrant) = Reduce(x);
        return quadrant % 2 == 0 ? Math.Tan(r) : -1 / Math.Tan(r);
    }

    // x = r + k pi/2 with |r| <= pi/4, by Excel's pi; the quadrant is k mod 4. The fused
    // multiply-add keeps x - k*PI()/2 exact where it cancels.
    private static (double Remainder, int Quadrant) Reduce(double x)
    {
        var k = Math.Round(x / (Math.PI / 2));
        var r = Math.FusedMultiplyAdd(-k, Math.PI / 2, x) - k * (PiLow / 2);
        return (r, (int)((long)k & 3));
    }

    // Excel computes ln(x + sqrt(x^2 + 1)) as written, which shows near 0: ASINH(1E-12) is
    // 1.000088900581841E-12. Far from 0 the exact function is used, where x^2 would overflow.
    private static CellValue Asinh(double x)
    {
        if (Math.Abs(x) > 1e8)
            return CellValue.Number(Math.Asinh(x));
        var magnitude = Math.Log(Math.Abs(x) + Math.Sqrt(x * x + 1));
        return CellValue.Number(x < 0 ? -magnitude : magnitude);
    }

    // Excel's ATAN2 takes x first: ATAN2(x, y) is the angle of the point (x, y).
    private static CellValue Atan2(double x, double y) =>
        x == 0 && y == 0 ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(Math.Atan2(y, x));
}
