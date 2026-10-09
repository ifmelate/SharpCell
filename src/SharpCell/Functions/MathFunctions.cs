using System;
using System.Collections.Generic;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Functions;

/// <summary>SUM-like aggregates and the arithmetic of the math category: powers, logarithms, factorials, divisors.</summary>
internal static class MathFunctions
{
    private static readonly ArgumentKind[] AnyArguments = [ArgumentKind.Any];
    private static readonly ArgumentKind[] ValueArgument = [ArgumentKind.Value];

    // Above 2^53 doubles are no longer consecutive integers; Excel's GCD and LCM refuse them.
    private const double MaxExactInteger = 9007199254740992;

    private const double MinNormal = 2.2250738585072014E-308;

    public static void Register(FunctionRegistry registry)
    {
        var max = FunctionRegistry.MaxArguments;
        registry.Add(new FunctionInfo("SUM", 1, max, AnyArguments, Sum));
        registry.Add(new FunctionInfo("AVERAGE", 1, max, AnyArguments, Average));
        registry.Add(new FunctionInfo("MIN", 1, max, AnyArguments, call => Extreme(call, Math.Min)));
        registry.Add(new FunctionInfo("MAX", 1, max, AnyArguments, call => Extreme(call, Math.Max)));
        registry.Add(new FunctionInfo("COUNT", 1, max, AnyArguments, Count));
        registry.Add(new FunctionInfo("COUNTA", 1, max, AnyArguments, CountA));
        registry.Add(new FunctionInfo("PRODUCT", 1, max, AnyArguments, Product));
        registry.Add(new FunctionInfo("SUMSQ", 1, max, AnyArguments, SumSq));
        registry.Add(new FunctionInfo("ABS", 1, 1, ValueArgument, Unary(x => Math.Abs(x))));
        registry.Add(new FunctionInfo("SIGN", 1, 1, ValueArgument, Unary(x => Math.Sign(x))));
        registry.Add(new FunctionInfo("SQRT", 1, 1, ValueArgument, Unary(x => x < 0 ? CellValue.Error(ErrorKind.Num) : Math.Sqrt(x))));
        registry.Add(new FunctionInfo("SQRTPI", 1, 1, ValueArgument, Unary(x => x < 0 ? CellValue.Error(ErrorKind.Num) : Math.Sqrt(x * Math.PI), toolPak: true)));
        registry.Add(new FunctionInfo("EXP", 1, 1, ValueArgument, Unary(Exp)));
        registry.Add(new FunctionInfo("LN", 1, 1, ValueArgument, Unary(x => x <= 0 ? CellValue.Error(ErrorKind.Num) : Math.Log(x))));
        registry.Add(new FunctionInfo("LOG10", 1, 1, ValueArgument, Unary(x => x <= 0 ? CellValue.Error(ErrorKind.Num) : Math.Log10(x))));
        registry.Add(new FunctionInfo("LOG", 1, 2, ValueArgument, Log));
        registry.Add(new FunctionInfo("POWER", 2, 2, ValueArgument, Power));
        registry.Add(new FunctionInfo("MOD", 2, 2, ValueArgument, Binary(Mod)));
        registry.Add(new FunctionInfo("QUOTIENT", 2, 2, ValueArgument, Binary(Quotient, toolPak: true)));
        registry.Add(new FunctionInfo("PI", 0, 0, ValueArgument, _ => CellValue.Number(Math.PI)));
        registry.Add(new FunctionInfo("FACT", 1, 1, ValueArgument, Unary(Fact)));
        registry.Add(new FunctionInfo("FACTDOUBLE", 1, 1, ValueArgument, Unary(FactDouble, toolPak: true)));
        registry.Add(new FunctionInfo("COMBIN", 2, 2, ValueArgument, Binary(Combin)));
        registry.Add(new FunctionInfo("COMBINA", 2, 2, ValueArgument, Binary(CombinA)));
        registry.Add(new FunctionInfo("MULTINOMIAL", 1, max, AnyArguments, Multinomial));
        registry.Add(new FunctionInfo("GCD", 1, max, AnyArguments, call => Divisors(call, Gcd)));
        registry.Add(new FunctionInfo("LCM", 1, max, AnyArguments, call => Divisors(call, Lcm)));
        registry.Add(new FunctionInfo("SERIESSUM", 4, 4, [ArgumentKind.Value, ArgumentKind.Value, ArgumentKind.Value, ArgumentKind.Any], SeriesSum));
        registry.Add(new FunctionInfo("RAND", 0, 0, ValueArgument, call => CellValue.Number(call.Context.Workbook.Random.NextDouble()))
        {
            IsVolatile = true,
        });
        registry.Add(new FunctionInfo("RANDBETWEEN", 2, 2, ValueArgument, RandBetween) { IsVolatile = true });
    }

    /// <summary>A body for a function of one number: the argument is coerced, an error passes through.</summary>
    internal static FunctionBody Unary(Func<double, CellValue> f, bool toolPak = false) => call =>
    {
        var x = toolPak ? ToolPakNumber(call, 0) : call.Number(0);
        return x.IsError ? x : f(x.AsNumber());
    };

    /// <summary>A body for a function of two numbers; the first error in argument order wins.</summary>
    internal static FunctionBody Binary(Func<double, double, CellValue> f, bool toolPak = false) => call =>
    {
        var x = toolPak ? ToolPakNumber(call, 0) : call.Number(0);
        if (x.IsError)
            return x;
        var y = toolPak ? ToolPakNumber(call, 1) : call.Number(1);
        return y.IsError ? y : f(x.AsNumber(), y.AsNumber());
    };

    /// <summary>
    /// A number argument of a function that came from the Analysis ToolPak (QUOTIENT, MROUND,
    /// SQRTPI and others): like any number argument, except that TRUE and FALSE are <c>#VALUE!</c>
    /// and an argument left empty is <c>#N/A</c>.
    /// </summary>
    internal static CellValue ToolPakNumber(FunctionCall call, int index)
    {
        var value = call.Value(index);
        return value.Kind switch
        {
            CellValueKind.Boolean => CellValue.Error(ErrorKind.Value),
            CellValueKind.Missing => CellValue.Error(ErrorKind.NA),
            _ => Coercion.ToNumber(value, call.Context.Culture, call.Context.DateSystem),
        };
    }

    private static Operand Sum(FunctionCall call)
    {
        var total = 0.0;
        var error = Aggregation.Numbers(call, n => total += n);
        return error.IsError ? error : CellValue.Number(total);
    }

    private static Operand Average(FunctionCall call)
    {
        var total = 0.0;
        var count = 0;
        var error = Aggregation.Numbers(call, n =>
        {
            total += n;
            count++;
        });
        if (error.IsError)
            return error;
        return count == 0 ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(total / count);
    }

    // MIN and MAX of no numbers is 0.
    private static Operand Extreme(FunctionCall call, Func<double, double, double> pick)
    {
        double? result = null;
        var error = Aggregation.Numbers(call, n => result = result is { } r ? pick(r, n) : n);
        return error.IsError ? error : CellValue.Number(result ?? 0);
    }

    // Ranges and arrays: numbers only. Direct arguments: anything that converts to a number.
    // Errors are never counted and never propagate.
    private static Operand Count(FunctionCall call)
    {
        var count = 0;
        Aggregation.ForEach(call, (value, source) =>
        {
            if (value.Kind == CellValueKind.Number)
                count++;
            else if (source == ValueSource.Direct && !value.IsError && value.Kind != CellValueKind.Empty
                     && !Coercion.ToNumber(value, call.Context.Culture, call.Context.DateSystem).IsError)
                count++;
            return true;
        });
        return CellValue.Number(count);
    }

    private static Operand CountA(FunctionCall call)
    {
        var count = 0;
        Aggregation.ForEach(call, (value, _) =>
        {
            if (value.Kind != CellValueKind.Empty)
                count++;
            return true;
        });
        return CellValue.Number(count);
    }

    // The product of no numbers is 0, not 1.
    private static Operand Product(FunctionCall call)
    {
        var product = 1.0;
        var any = false;
        var error = Aggregation.Numbers(call, n =>
        {
            product *= n;
            any = true;
        });
        return error.IsError ? error : CellValue.Number(any ? product : 0);
    }

    private static Operand SumSq(FunctionCall call)
    {
        var total = 0.0;
        var error = Aggregation.Numbers(call, n => total += n * n);
        return error.IsError ? error : CellValue.Number(total);
    }

    private static Operand Log(FunctionCall call)
    {
        var x = call.Number(0);
        if (x.IsError)
            return x;
        var b = call.Number(1, 10);
        if (b.IsError)
            return b;

        double number = x.AsNumber(), @base = b.AsNumber();
        if (number <= 0 || @base <= 0)
            return CellValue.Error(ErrorKind.Num);
        if (@base == 1)
            return CellValue.Error(ErrorKind.Div0);
        return CellValue.Number(@base == 10 ? Math.Log10(number) : Math.Log(number) / Math.Log(@base));
    }

    // Excel has no subnormal numbers: EXP(-745) is 0, not 5E-324.
    private static CellValue Exp(double x)
    {
        var result = Math.Exp(x);
        return CellValue.Number(result < MinNormal ? 0 : result);
    }

    // POWER is the ^ operator.
    private static Operand Power(FunctionCall call)
    {
        var x = call.Number(0);
        if (x.IsError)
            return x;
        var y = call.Number(1);
        if (y.IsError)
            return y;
        return Operators.Binary(BinaryOperator.Power, x, y, call.Context.Culture, call.Context.DateSystem, last: false);
    }

    // The result takes the divisor's sign: MOD(-3,2) is 1. The remainder is exact, as Excel's is:
    // MOD(5,1E-10) is 9.99998E-11 because 1E-10 is slightly more than 10^-10 in binary. Excel
    // refuses quotients too large to leave a meaningful remainder.
    private static CellValue Mod(double n, double d)
    {
        if (d == 0)
            return CellValue.Error(ErrorKind.Div0);
        if (Math.Abs(n / d) >= MaxExactInteger)
            return CellValue.Error(ErrorKind.Num);
        var remainder = n % d;
        return CellValue.Number(remainder != 0 && remainder < 0 != d < 0 ? remainder + d : remainder);
    }

    private static CellValue Quotient(double n, double d) =>
        d == 0 ? CellValue.Error(ErrorKind.Div0) : CellValue.Number(Math.Truncate(n / d));

    // 170! is the largest factorial a double holds.
    private static CellValue Fact(double x)
    {
        if (x < 0)
            return CellValue.Error(ErrorKind.Num);
        var n = Math.Truncate(x);
        if (n > 170)
            return CellValue.Error(ErrorKind.Num);
        var result = 1.0;
        for (var i = 2; i <= n; i++)
            result *= i;
        return CellValue.Number(result);
    }

    // n!! = n(n-2)(n-4)...; 0!! and (-1)!! are 1.
    private static CellValue FactDouble(double x)
    {
        if (x < -1)
            return CellValue.Error(ErrorKind.Num);
        var n = Math.Truncate(x);
        if (n > 300)
            return CellValue.Error(ErrorKind.Num);
        var result = 1.0;
        for (var i = n; i > 1; i -= 2)
            result *= i;
        return CellValue.Number(result);
    }

    private static CellValue Combin(double x, double y)
    {
        double n = Math.Truncate(x), k = Math.Truncate(y);
        if (n < 0 || k < 0 || k > n)
            return CellValue.Error(ErrorKind.Num);
        return CellValue.Number(Choose(n, k));
    }

    // Combinations with repetitions: COMBIN(n+k-1, k).
    private static CellValue CombinA(double x, double y)
    {
        double n = Math.Truncate(x), k = Math.Truncate(y);
        if (n < 0 || k < 0 || (n == 0 && k > 0))
            return CellValue.Error(ErrorKind.Num);
        return n == 0 ? CellValue.Number(1) : CellValue.Number(Choose(n + k - 1, k));
    }

    // n choose k by the multiplicative formula over the smaller side; overflow gives infinity, which is #NUM!.
    private static double Choose(double n, double k)
    {
        k = Math.Min(k, n - k);
        var result = 1.0;
        for (var i = 1.0; i <= k && double.IsFinite(result); i++)
            result = result * (n - k + i) / i;
        return Math.Round(result) is var rounded && Math.Abs(rounded - result) <= Math.Abs(result) * 1e-12 ? rounded : result;
    }

    private static Operand Multinomial(FunctionCall call)
    {
        var terms = new List<double>();
        var error = ReadIntegers(call, terms);
        if (error is { } e)
            return e;
        if (terms.Count == 0)
            return CellValue.Error(ErrorKind.Value);

        // (a+b+c)! / (a! b! c!) as a product of binomial coefficients, which overflows only when the result does.
        var result = 1.0;
        var total = 0.0;
        foreach (var term in terms)
        {
            total += term;
            result *= Choose(total, term);
        }

        return CellValue.Number(result);
    }

    private static Operand Divisors(FunctionCall call, Func<double, double, double> combine)
    {
        var values = new List<double>();
        var error = ReadIntegers(call, values);
        if (error is { } e)
            return e;
        if (values.Count == 0)
            return CellValue.Error(ErrorKind.Value);

        var result = values[0];
        for (var i = 1; i < values.Count; i++)
        {
            result = combine(result, values[i]);
            if (result > MaxExactInteger)
                return CellValue.Error(ErrorKind.Num);
        }

        return CellValue.Number(result);
    }

    private static double Gcd(double a, double b)
    {
        while (b != 0)
            (a, b) = (b, a % b);
        return a;
    }

    private static double Lcm(double a, double b) => a == 0 || b == 0 ? 0 : a / Gcd(a, b) * b;

    /// <summary>
    /// The non-negative integers of GCD, LCM and MULTINOMIAL. Every value, from ranges and arrays
    /// too, is coerced from text and truncated. An empty cell referenced on its own is skipped, but
    /// inside a larger range it counts as 0. Booleans are <c>#VALUE!</c>, an omitted argument
    /// <c>#N/A</c>, negatives and values from 2^53 up <c>#NUM!</c>.
    /// </summary>
    private static CellValue? ReadIntegers(FunctionCall call, List<double> values)
    {
        var context = call.Context;
        CellValue? error = null;
        var visited = 0;

        // After an error the remaining cells are still read when a dirty cell was met, so that
        // every dirty input is collected in one pass (see Aggregation.ForEach).
        bool Stop() => error is not null && !call.MetPendingInput;

        void Add(CellValue value)
        {
            if (error is not null)
                return;
            var number = value.Kind switch
            {
                CellValueKind.Boolean => CellValue.Error(ErrorKind.Value),
                CellValueKind.Empty => CellValue.Number(0),
                _ => Coercion.ToNumber(value, context.Culture, context.DateSystem),
            };
            if (!number.IsError && (number.AsNumber() < 0 || number.AsNumber() >= MaxExactInteger))
                number = CellValue.Error(ErrorKind.Num);
            if (number.IsError)
                error = number;
            else
                values.Add(Math.Truncate(number.AsNumber()));
        }

        for (var i = 0; i < call.Count && !Stop(); i++)
        {
            if (call.IsMissing(i))
            {
                Add(CellValue.Error(ErrorKind.NA));
                continue;
            }

            var argument = call[i];
            if (argument.Reference is not { } reference)
            {
                var value = argument.Value;
                if (value.Kind != CellValueKind.Array)
                {
                    Add(value);
                    continue;
                }

                foreach (var element in value.AsArray())
                    Add(element);
                continue;
            }

            foreach (var (sheet, area) in reference.Areas)
            {
                if (area.IsSingleCell)
                {
                    var single = context.ReadCell(sheet, area.FirstRow, area.FirstColumn);
                    if (single.Kind != CellValueKind.Empty)
                        Add(single);
                    continue;
                }

                long stored = 0;
                foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
                {
                    if (++visited % 4096 == 0)
                        context.CancellationToken.ThrowIfCancellationRequested();
                    stored++;
                    Add(context.ReadCell(sheet, cell));
                    if (Stop())
                        return error;
                }

                if (stored < area.CellCount)
                    Add(CellValue.Empty);
            }
        }

        return error;
    }

    // x^n * (a1 + a2 x^m + a3 x^2m + ...). Coefficients must be numbers; empty cells count as 0.
    private static Operand SeriesSum(FunctionCall call)
    {
        var x = ToolPakNumber(call, 0);
        if (x.IsError)
            return x;
        var n = ToolPakNumber(call, 1);
        if (n.IsError)
            return n;
        var m = ToolPakNumber(call, 2);
        if (m.IsError)
            return m;
        var argument = call.Value(3);
        if (argument.IsError)
            return argument;

        double total = 0, power = n.AsNumber(), step = m.AsNumber(), @base = x.AsNumber();
        foreach (var coefficient in argument.Kind == CellValueKind.Array ? argument.AsArray() : new[,] { { argument } })
        {
            if (coefficient.IsError)
                return coefficient;
            if (coefficient.Kind is not (CellValueKind.Number or CellValueKind.Empty))
                return CellValue.Error(ErrorKind.Value);

            // 0^0 is #NUM!, as with the ^ operator.
            if (@base == 0 && power == 0)
                return CellValue.Error(ErrorKind.Num);
            if (coefficient.Kind == CellValueKind.Number)
                total += coefficient.AsNumber() * Math.Pow(@base, power);
            power += step;
        }

        return CellValue.Number(total);
    }

    // A whole number in [CEILING(bottom), FLOOR(top)].
    private static Operand RandBetween(FunctionCall call)
    {
        var bottom = ToolPakNumber(call, 0);
        if (bottom.IsError)
            return bottom;
        var top = ToolPakNumber(call, 1);
        if (top.IsError)
            return top;

        double low = Math.Ceiling(bottom.AsNumber()), high = Math.Floor(top.AsNumber());
        if (low > high)
            return CellValue.Error(ErrorKind.Num);
        var random = call.Context.Workbook.Random.NextDouble();
        return CellValue.Number(Math.Min(high, low + Math.Floor(random * (high - low + 1))));
    }
}
