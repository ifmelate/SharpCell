using SharpCell.Evaluation;

namespace SharpCell.Functions;

internal static class InformationFunctions
{
    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("ISBLANK", 1, 1, [ArgumentKind.Value], call => CellValue.Boolean(call.Value(0).Kind == CellValueKind.Empty)));
        registry.Add(new FunctionInfo("ISERROR", 1, 1, [ArgumentKind.Value], call => CellValue.Boolean(call.Value(0).IsError)));
        registry.Add(new FunctionInfo("ROWS", 1, 1, [ArgumentKind.Any], call => Dimension(call, rows: true)));
        registry.Add(new FunctionInfo("COLUMNS", 1, 1, [ArgumentKind.Any], call => Dimension(call, rows: false)));
    }

    private static Operand Dimension(FunctionCall call, bool rows)
    {
        var argument = call[0];
        if (argument.Reference is { } reference)
        {
            if (!reference.IsSingleArea)
                return CellValue.Error(ErrorKind.Ref);
            var area = reference.Areas[0].Area;
            return CellValue.Number(rows ? area.Rows : area.Columns);
        }

        var value = argument.Value;
        if (value.IsError)
            return value;
        if (value.Kind != CellValueKind.Array)
            return CellValue.Number(1);
        return CellValue.Number(value.AsArray().GetLength(rows ? 0 : 1));
    }
}
