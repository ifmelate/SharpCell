using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>Argument reading shared by the financial functions.</summary>
internal static class FinancialArguments
{
    /// <summary>Marks a required argument in the defaults of <see cref="TryNumbers"/>.</summary>
    public const double Required = double.NaN;

    /// <summary>
    /// Reads the arguments as numbers, left to right; the first error wins. <paramref name="defaults"/>
    /// has one entry per parameter: <see cref="Required"/>, or the value an optional argument takes
    /// when it is absent or left empty.
    /// </summary>
    public static bool TryNumbers(FunctionCall call, out double[] values, out CellValue error, params double[] defaults) =>
        TryRead(call, rejectLogical: false, out values, out error, defaults);

    /// <summary>
    /// As <see cref="TryNumbers"/>, for the functions that came from the Analysis ToolPak add-in:
    /// they take numbers and number text but refuse TRUE and FALSE with <c>#VALUE!</c>, typed or
    /// read from a cell.
    /// </summary>
    public static bool TryAddInNumbers(FunctionCall call, out double[] values, out CellValue error, params double[] defaults) =>
        TryRead(call, rejectLogical: true, out values, out error, defaults);

    /// <summary>One argument as <see cref="TryAddInNumbers"/> reads it.</summary>
    public static bool TryAddInNumber(FunctionCall call, int index, double absent, out double value, out CellValue error) =>
        TryRead(call, index, absent, rejectLogical: true, out value, out error);

    private static bool TryRead(FunctionCall call, bool rejectLogical, out double[] values, out CellValue error, double[] defaults)
    {
        values = new double[defaults.Length];
        for (var i = 0; i < defaults.Length; i++)
        {
            if (!TryRead(call, i, defaults[i], rejectLogical, out values[i], out error))
                return false;
        }

        error = default;
        return true;
    }

    private static bool TryRead(FunctionCall call, int index, double absent, bool rejectLogical, out double result, out CellValue error)
    {
        result = 0;
        error = default;
        if (!double.IsNaN(absent) && !call.Has(index))
        {
            result = absent;
            return true;
        }

        var value = call.Value(index);
        if (rejectLogical && value.Kind == CellValueKind.Boolean)
        {
            error = CellValue.Error(ErrorKind.Value);
            return false;
        }

        var number = Coercion.ToNumber(value, call.Context.Culture, call.Context.DateSystem);
        if (number.IsError)
        {
            error = number;
            return false;
        }

        result = number.AsNumber();
        return true;
    }

    /// <summary>A date argument: the serial number's whole part, which must be a date Excel knows.</summary>
    public static bool TryDate(double serial, DateSystem system, out FinancialDate date)
    {
        date = default;
        if (!DateSerial.TryToDate(serial, system, out var year, out var month, out var day) || day == 0)
            return false;
        date = new FinancialDate(year, month, day);
        return true;
    }
}
