using SharpCell.Functions;

namespace SharpCell.Xlsx;

/// <summary>Which names in a formula stand for functions passed as values, as the evaluator resolves them.</summary>
internal static class FunctionValues
{
    /// <summary>
    /// Whether a name without a sheet means a function: not a defined name visible from the sheet or a
    /// table, and a function SharpCell has (built in or custom), other than LET and LAMBDA.
    /// </summary>
    public static bool IsFunction(Workbook workbook, Worksheet? sheet, string name)
    {
        var upper = name.ToUpperInvariant();
        if (upper is "LET" or "LAMBDA"
            || (sheet is not null && workbook.Names.TryGet(upper, sheet, out _))
            || workbook.Names.TryGet(upper, null, out _)
            || workbook.TryGetTable(upper, out _))
            return false;
        return workbook.TryGetFunction(upper, out var info) && info!.Status != FunctionStatus.NotImplemented;
    }
}
