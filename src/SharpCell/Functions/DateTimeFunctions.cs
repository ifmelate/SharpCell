using SharpCell.Evaluation;

namespace SharpCell.Functions;

internal static class DateTimeFunctions
{
    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("NOW", 0, 0, [ArgumentKind.Value], Now) { IsVolatile = true });
    }

    private static Operand Now(FunctionCall call)
    {
        var workbook = call.Context.Workbook;
        return CellValue.Number(DateSerial.FromDateTime(workbook.Clock.GetLocalNow().DateTime, workbook.DateSystem));
    }
}
