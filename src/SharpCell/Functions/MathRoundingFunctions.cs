using System;
using System.Globalization;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>ROUND and its relatives, multiples (MROUND, CEILING*, FLOOR*), INT, TRUNC, EVEN and ODD.</summary>
internal static class MathRoundingFunctions
{
    private static readonly ArgumentKind[] ValueArguments = [ArgumentKind.Value];

    private enum Rounding
    {
        Nearest,
        Up,
        Down,
    }

    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("ROUND", 2, 2, ValueArguments, MathFunctions.Binary((x, d) => Round(x, d, Rounding.Nearest))));
        registry.Add(new FunctionInfo("ROUNDUP", 2, 2, ValueArguments, MathFunctions.Binary((x, d) => Round(x, d, Rounding.Up))));
        registry.Add(new FunctionInfo("ROUNDDOWN", 2, 2, ValueArguments, MathFunctions.Binary((x, d) => Round(x, d, Rounding.Down))));
        registry.Add(new FunctionInfo("TRUNC", 1, 2, ValueArguments, Trunc));
        registry.Add(new FunctionInfo("INT", 1, 1, ValueArguments, MathFunctions.Unary(x => Math.Floor(x))));
        registry.Add(new FunctionInfo("EVEN", 1, 1, ValueArguments, MathFunctions.Unary(Even)));
        registry.Add(new FunctionInfo("ODD", 1, 1, ValueArguments, MathFunctions.Unary(Odd)));
        registry.Add(new FunctionInfo("MROUND", 2, 2, ValueArguments, MathFunctions.Binary(MRound, toolPak: true)));
        registry.Add(new FunctionInfo("CEILING", 2, 2, ValueArguments, MathFunctions.Binary(Ceiling)));
        registry.Add(new FunctionInfo("FLOOR", 2, 2, ValueArguments, MathFunctions.Binary(Floor)));
        registry.Add(new FunctionInfo("CEILING.MATH", 1, 3, ValueArguments, call => MathMultiple(call, ceiling: true)));
        registry.Add(new FunctionInfo("FLOOR.MATH", 1, 3, ValueArguments, call => MathMultiple(call, ceiling: false)));
        registry.Add(new FunctionInfo("CEILING.PRECISE", 1, 2, ValueArguments, call => PreciseMultiple(call, ceiling: true)));
        registry.Add(new FunctionInfo("ISO.CEILING", 1, 2, ValueArguments, call => PreciseMultiple(call, ceiling: true)));
        registry.Add(new FunctionInfo("FLOOR.PRECISE", 1, 2, ValueArguments, call => PreciseMultiple(call, ceiling: false)));
    }

    /// <summary>
    /// A number at 15 significant digits, the precision Excel works in. Rounding first to these
    /// digits keeps binary noise from crossing a rounding boundary: 2.675 is stored as
    /// 2.67499999999999982..., and Excel's ROUND(2.675, 2) is still 2.68.
    /// </summary>
    internal static double Significant(double value) =>
        value == 0 || !double.IsFinite(value)
            ? value
            : double.Parse(value.ToString("E14", CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture);

    // Digits are truncated (ROUND(x, 1.9) rounds at one decimal); negative digits round to tens, hundreds...
    private static CellValue Round(double x, double digits, Rounding mode)
    {
        digits = Math.Truncate(digits);
        if (x == 0)
            return CellValue.Number(0);

        // Beyond 10^308 every finite number is less than one unit.
        if (digits < -308)
            return mode == Rounding.Up ? CellValue.Error(ErrorKind.Num) : CellValue.Number(0);

        // Multiplying by a power of ten below 10^-22 is inexact, dividing by its inverse is not.
        var scale = Math.Pow(10, Math.Abs(digits));
        var scaled = digits >= 0 ? x * scale : x / scale;

        // A position past the digits the number has leaves it unchanged.
        if (!double.IsFinite(scaled) || Math.Abs(scaled) >= 1e15)
            return CellValue.Number(x);

        scaled = Significant(scaled);
        var rounded = mode switch
        {
            Rounding.Nearest => Math.Round(scaled, MidpointRounding.AwayFromZero),
            Rounding.Up => scaled > 0 ? Math.Ceiling(scaled) : Math.Floor(scaled),
            _ => Math.Truncate(scaled),
        };
        return CellValue.Number(digits >= 0 ? rounded / scale : rounded * scale);
    }

    private static Operand Trunc(FunctionCall call)
    {
        var x = call.Number(0);
        if (x.IsError)
            return x;
        var digits = call.Number(1, 0);
        return digits.IsError ? digits : Round(x.AsNumber(), digits.AsNumber(), Rounding.Down);
    }

    // Away from zero to an even integer: EVEN(1.5) is 2, EVEN(-1) is -2.
    private static CellValue Even(double x)
    {
        var magnitude = Math.Ceiling(Math.Abs(x) / 2) * 2;
        return CellValue.Number(x < 0 ? -magnitude : magnitude);
    }

    // Away from zero to an odd integer; ODD(0) is 1.
    private static CellValue Odd(double x)
    {
        var magnitude = Math.Ceiling(Math.Abs(x));
        if (magnitude % 2 == 0)
            magnitude++;
        return CellValue.Number(x < 0 ? -magnitude : magnitude);
    }

    // To the nearest multiple, halves away from zero. Number and multiple must not have opposite signs.
    private static CellValue MRound(double x, double multiple)
    {
        if (multiple == 0)
            return CellValue.Number(0);
        if ((x > 0 && multiple < 0) || (x < 0 && multiple > 0))
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(Math.Round(Ratio(x, multiple), MidpointRounding.AwayFromZero) * multiple);
    }

    // CEILING and FLOOR round x/significance up or down. A positive number with a negative
    // significance is #NUM!; a negative number with a positive one rounds toward zero (CEILING)
    // or away from it (FLOOR), with a negative one the other way round.
    private static CellValue Ceiling(double x, double significance)
    {
        if (significance == 0)
            return CellValue.Number(0);
        if (x > 0 && significance < 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(Math.Ceiling(Ratio(x, significance)) * significance);
    }

    private static CellValue Floor(double x, double significance)
    {
        if (significance == 0)
            return x == 0 ? CellValue.Number(0) : CellValue.Error(ErrorKind.Div0);
        if (x > 0 && significance < 0)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(Math.Floor(Ratio(x, significance)) * significance);
    }

    // CEILING.MATH and FLOOR.MATH: the sign of the significance is ignored; a nonzero mode makes
    // negative numbers round the other way (CEILING.MATH away from zero, FLOOR.MATH toward it).
    private static Operand MathMultiple(FunctionCall call, bool ceiling)
    {
        var x = call.Number(0);
        if (x.IsError)
            return x;
        var significance = call.Number(1, 1);
        if (significance.IsError)
            return significance;
        var mode = call.Number(2, 0);
        if (mode.IsError)
            return mode;

        var step = Math.Abs(significance.AsNumber());
        if (step == 0)
            return CellValue.Number(0);
        var number = x.AsNumber();
        var up = number < 0 && mode.AsNumber() != 0 ? !ceiling : ceiling;
        var ratio = Ratio(number, step);
        return CellValue.Number((up ? Math.Ceiling(ratio) : Math.Floor(ratio)) * step);
    }

    // CEILING.PRECISE, ISO.CEILING and FLOOR.PRECISE: the sign of the significance is ignored.
    private static Operand PreciseMultiple(FunctionCall call, bool ceiling)
    {
        var x = call.Number(0);
        if (x.IsError)
            return x;
        var significance = call.Number(1, 1);
        if (significance.IsError)
            return significance;

        var step = Math.Abs(significance.AsNumber());
        if (step == 0)
            return CellValue.Number(0);
        var ratio = Ratio(x.AsNumber(), step);
        return CellValue.Number((ceiling ? Math.Ceiling(ratio) : Math.Floor(ratio)) * step);
    }

    // The number of multiples, at Excel's precision: 0.3/0.1 is 2.9999999999999996 in binary and 3 here.
    private static double Ratio(double x, double multiple) => Significant(x / multiple);
}
