using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Functions;

internal static class LogicalFunctions
{
    public static void Register(FunctionRegistry registry)
    {
        var max = FunctionRegistry.MaxArguments;
        registry.Add(new FunctionInfo("IF", 2, 3, [ArgumentKind.ScalarAny, ArgumentKind.Lazy], If));
        registry.Add(new FunctionInfo("IFERROR", 2, 2, [ArgumentKind.ScalarAny, ArgumentKind.Lazy], IfError));
        registry.Add(new FunctionInfo("IFNA", 2, 2, [ArgumentKind.ScalarAny, ArgumentKind.Lazy], IfNa));
        registry.Add(new FunctionInfo("IFS", 2, max, [ArgumentKind.Lazy], Ifs));
        registry.Add(new FunctionInfo("SWITCH", 3, max, [ArgumentKind.ScalarAny, ArgumentKind.Lazy], Switch));
        registry.Add(new FunctionInfo("CHOOSE", 2, max, [ArgumentKind.Value, ArgumentKind.Lazy], Choose));
        registry.Add(new FunctionInfo("AND", 1, max, [ArgumentKind.Any], call => Junction(call, Join.And)));
        registry.Add(new FunctionInfo("OR", 1, max, [ArgumentKind.Any], call => Junction(call, Join.Or)));
        registry.Add(new FunctionInfo("XOR", 1, max, [ArgumentKind.Any], call => Junction(call, Join.Xor)));
        registry.Add(new FunctionInfo("NOT", 1, 1, [ArgumentKind.Value], Not));
        registry.Add(new FunctionInfo("TRUE", 0, 0, [], _ => CellValue.True));
        registry.Add(new FunctionInfo("FALSE", 0, 0, [], _ => CellValue.False));
    }

    private enum Join
    {
        And,
        Or,
        Xor,
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

    // IFERROR for #N/A only: other errors pass through.
    private static Operand IfNa(FunctionCall call)
    {
        var value = call.Value(0);
        if (value.Kind == CellValueKind.Array)
        {
            var fallback = Branch(call, 1, absent: CellValue.Number(0));
            return ArrayMath.Map(value, fallback, (v, f) => IsNa(v) ? f : v);
        }

        return IsNa(value) ? BranchOperand(call, 1, CellValue.Number(0)) : value;
    }

    private static bool IsNa(CellValue value) => value.IsError && value.AsError() == ErrorKind.NA;

    // IFS(condition1, value1, ...): the value of the first true condition, #N/A when none is.
    // Conditions are evaluated in order and only as far as needed, values only when chosen. An
    // array condition decides element-wise, so from there on every pair is needed.
    private static Operand Ifs(FunctionCall call)
    {
        // Excel refuses an unpaired condition when the formula is entered; files cannot hold one.
        if (call.Count % 2 != 0)
            return CellValue.Error(ErrorKind.Value);

        for (var i = 0; i < call.Count; i += 2)
        {
            var condition = ScalarArgument(call, i);
            if (condition.Kind == CellValueKind.Array)
                return IfsElementWise(call, i, condition);

            var test = Coercion.ToBoolean(condition);
            if (test.IsError)
                return test;
            if (test.AsBoolean())
                return BranchOperand(call, i + 1, CellValue.Number(0));
        }

        return CellValue.Error(ErrorKind.NA);
    }

    // Folds the remaining pairs from the last one back: each condition picks its value or what
    // the later pairs give, so the first true condition wins at every position.
    private static CellValue IfsElementWise(FunctionCall call, int first, CellValue firstCondition)
    {
        CellValue result = CellValue.Error(ErrorKind.NA);
        for (var i = call.Count - 2; i >= first; i -= 2)
        {
            var condition = i == first ? firstCondition : ScalarArgument(call, i);
            var value = Branch(call, i + 1, absent: CellValue.Number(0));
            result = ArrayMath.Map(condition, value, result, (c, v, rest) =>
            {
                var test = Coercion.ToBoolean(c);
                return test.IsError ? test : test.AsBoolean() ? v : rest;
            });
        }

        return result;
    }

    // SWITCH(expression, value1, result1, ..., [default]): compares with = (text ignoring case);
    // an error in the expression or in a compared value is the result. Values are evaluated in
    // order until one matches; no match and no default is #N/A.
    private static Operand Switch(FunctionCall call)
    {
        var expression = call.Value(0);
        var pairs = (call.Count - 1) / 2;
        var hasDefault = (call.Count - 1) % 2 == 1;
        if (expression.Kind == CellValueKind.Array)
            return SwitchElementWise(call, expression, pairs, hasDefault);
        if (expression.IsError)
            return expression;

        for (var p = 0; p < pairs; p++)
        {
            var index = 1 + 2 * p;
            var candidate = ScalarArgument(call, index);
            if (candidate.Kind == CellValueKind.Array)
                return SwitchElementWise(call, expression, pairs, hasDefault);

            var equal = SwitchEquals(expression, candidate, call.Context);
            if (equal.IsError)
                return equal;
            if (equal.AsBoolean())
                return BranchOperand(call, index + 1, CellValue.Number(0));
        }

        return hasDefault ? BranchOperand(call, call.Count - 1, CellValue.Number(0)) : CellValue.Error(ErrorKind.NA);
    }

    private static CellValue SwitchElementWise(FunctionCall call, CellValue expression, int pairs, bool hasDefault)
    {
        CellValue result = hasDefault ? Branch(call, call.Count - 1, CellValue.Number(0)) : CellValue.Error(ErrorKind.NA);
        for (var p = pairs - 1; p >= 0; p--)
        {
            var index = 1 + 2 * p;
            var candidate = ScalarArgument(call, index);
            var value = Branch(call, index + 1, CellValue.Number(0));
            var test = ArrayMath.Map(expression, candidate, (e, c) => e.IsError ? e : SwitchEquals(e, c, call.Context));
            result = ArrayMath.Map(test, value, result, (t, v, rest) => t.IsError ? t : t.AsBoolean() ? v : rest);
        }

        return result;
    }

    private static CellValue SwitchEquals(CellValue expression, CellValue candidate, EvaluationContext context) =>
        Operators.Binary(BinaryOperator.Equal, expression, candidate, context.Culture, context.DateSystem, last: false);

    // A lazy argument read the way a ScalarAny one is: in a formula without the dynamic array
    // flag a range is reduced by implicit intersection first.
    private static CellValue ScalarArgument(FunctionCall call, int index)
    {
        if (call.IsMissing(index))
            return CellValue.Empty;

        var context = call.Context;
        var outer = context.Legacy;
        context.Legacy = context.LegacyFormula;
        try
        {
            return Evaluator.ToValue(Evaluator.LegacyScalar(call[index], context), context);
        }
        finally
        {
            context.Legacy = outer;
        }
    }

    private static Operand Choose(FunctionCall call)
    {
        var index = Coercion.ToNumber(call.Value(0), call.Context.Culture, call.Context.DateSystem);
        if (index.IsError)
            return index;

        var choice = (long)System.Math.Truncate(index.AsNumber());
        if (choice < 1 || choice >= call.Count)
            return CellValue.Error(ErrorKind.Value);
        return BranchOperand(call, (int)choice, CellValue.Number(0));
    }

    // Ranges and arrays: booleans and numbers count, text and empty cells are skipped. Direct text
    // must read as TRUE or FALSE. With nothing to judge the result is #VALUE!.
    // XOR is true when an odd number of the values are true.
    private static Operand Junction(FunctionCall call, Join join)
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

            // Text typed into the call counts only if it reads as TRUE or FALSE; other text is skipped.
            var flag = Coercion.ToBoolean(value);
            if (flag.IsError && value.Kind == CellValueKind.Text)
                return true;
            if (flag.IsError)
            {
                error = flag;
                return false;
            }

            result = join switch
            {
                Join.And => (result ?? true) && flag.AsBoolean(),
                Join.Or => (result ?? false) || flag.AsBoolean(),
                _ => (result ?? false) ^ flag.AsBoolean(),
            };
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
