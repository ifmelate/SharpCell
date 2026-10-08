using SharpCell.Evaluation;

namespace SharpCell.Functions;

internal static class LogicalFunctions
{
    public static void Register(FunctionRegistry registry)
    {
        var max = FunctionRegistry.MaxArguments;
        registry.Add(new FunctionInfo("IF", 2, 3, [ArgumentKind.Any, ArgumentKind.Lazy], If));
        registry.Add(new FunctionInfo("IFERROR", 2, 2, [ArgumentKind.Any, ArgumentKind.Lazy], IfError));
        registry.Add(new FunctionInfo("CHOOSE", 2, max, [ArgumentKind.Value, ArgumentKind.Lazy], Choose));
        registry.Add(new FunctionInfo("AND", 1, max, [ArgumentKind.Any], call => Junction(call, isAnd: true)));
        registry.Add(new FunctionInfo("OR", 1, max, [ArgumentKind.Any], call => Junction(call, isAnd: false)));
        registry.Add(new FunctionInfo("NOT", 1, 1, [ArgumentKind.Value], Not));
    }

    // A scalar condition evaluates only the chosen branch. An array condition picks element-wise,
    // so both branches are needed: IF({TRUE,FALSE},{1,2},{3,4}) is {1,4}.
    private static Operand If(FunctionCall call)
    {
        var condition = call.Value(0);
        if (condition.Kind == CellValueKind.Array)
        {
            var whenTrue = Branch(call, 1, absent: CellValue.True);
            var whenFalse = Branch(call, 2, absent: CellValue.False);
            return ArrayMath.Map(condition, whenTrue, whenFalse, (c, t, f) =>
            {
                var test = Coercion.ToBoolean(c);
                return test.IsError ? test : test.AsBoolean() ? t : f;
            });
        }

        var test = Coercion.ToBoolean(condition);
        if (test.IsError)
            return test;
        return test.AsBoolean() ? BranchOperand(call, 1, CellValue.True) : BranchOperand(call, 2, CellValue.False);
    }

    private static Operand IfError(FunctionCall call)
    {
        var value = call.Value(0);
        if (value.Kind == CellValueKind.Array)
        {
            var fallback = Branch(call, 1, absent: CellValue.Number(0));
            return ArrayMath.Map(value, fallback, (v, f) => v.IsError ? f : v);
        }

        return value.IsError ? BranchOperand(call, 1, CellValue.Number(0)) : value;
    }

    private static Operand Choose(FunctionCall call)
    {
        var index = Coercion.ToNumber(call.Value(0), call.Context.Culture);
        if (index.IsError)
            return index;

        var choice = (long)System.Math.Truncate(index.AsNumber());
        if (choice < 1 || choice >= call.Count)
            return CellValue.Error(ErrorKind.Value);
        return BranchOperand(call, (int)choice, CellValue.Number(0));
    }

    // Ranges and arrays: booleans and numbers count, text and empty cells are skipped. Direct text
    // must read as TRUE or FALSE. With nothing to judge the result is #VALUE!.
    private static Operand Junction(FunctionCall call, bool isAnd)
    {
        bool? result = null;
        CellValue? error = null;
        Aggregation.ForEach(call, (value, source) =>
        {
            if (value.IsError)
            {
                error = value;
                return false;
            }

            if (value.Kind == CellValueKind.Empty || (value.Kind == CellValueKind.Text && source != ValueSource.Direct))
                return true;

            var flag = Coercion.ToBoolean(value);
            if (flag.IsError)
            {
                error = flag;
                return false;
            }

            result = isAnd ? (result ?? true) && flag.AsBoolean() : (result ?? false) || flag.AsBoolean();
            return true;
        });

        if (error is { } e)
            return e;
        return result is { } r ? CellValue.Boolean(r) : CellValue.Error(ErrorKind.Value);
    }

    private static Operand Not(FunctionCall call)
    {
        var flag = Coercion.ToBoolean(call.Value(0));
        return flag.IsError ? flag : CellValue.Boolean(!flag.AsBoolean());
    }

    // An omitted branch (IF(FALSE,1,)) gives 0; an absent one (IF(FALSE,1)) gives the default.
    private static Operand BranchOperand(FunctionCall call, int index, CellValue absent)
    {
        if (index >= call.Count)
            return absent;
        return call.IsMissing(index) ? CellValue.Number(0) : call[index];
    }

    private static CellValue Branch(FunctionCall call, int index, CellValue absent) =>
        Evaluator.ToValue(BranchOperand(call, index, absent), call.Context);
}
