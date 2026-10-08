using System;
using System.Collections.Generic;
using SharpCell.Parsing;

namespace SharpCell.Evaluation;

/// <summary>LET, LAMBDA and calls of function values.</summary>
internal static class Lambdas
{
    /// <summary>Nested lambda calls deeper than this give #NUM! (spec: recursion is allowed but bounded).</summary>
    public const int MaxCallDepth = 1024;

    /// <summary>LET and LAMBDA are syntax, not registry functions; returns false for other names.</summary>
    public static bool TryEvaluateSpecialForm(FunctionNode node, EvaluationContext context, out Operand result)
    {
        switch (node.Name)
        {
            case "LET":
                result = Let(node.Arguments, context);
                return true;
            case "LAMBDA":
                result = Lambda(node.Arguments, context);
                return true;
            default:
                result = default;
                return false;
        }
    }

    /// <summary>Calls a value with argument nodes evaluated in the caller's scope.</summary>
    public static Operand Call(CellValue callee, IReadOnlyList<FormulaNode> argumentNodes, EvaluationContext context)
    {
        if (callee.IsError)
            return callee;
        if (callee.Kind != CellValueKind.Lambda || callee.AsLambda() is not Closure closure)
            return CellValue.Error(ErrorKind.Value);

        var pendingBefore = context.Pending.Count;
        var arguments = new Operand[argumentNodes.Count];
        for (var i = 0; i < arguments.Length; i++)
            arguments[i] = Evaluator.Evaluate(argumentNodes[i], context);

        // An argument met a dirty cell: the body could take a branch the real value would not.
        if (context.Pending.Count > pendingBefore)
            return EvaluationContext.PendingPlaceholder;

        return Invoke(closure, arguments, context);
    }

    public static Operand Invoke(Closure closure, IReadOnlyList<Operand> arguments, EvaluationContext context)
    {
        if (arguments.Count > closure.Parameters.Length || arguments.Count < closure.RequiredCount)
            return CellValue.Error(ErrorKind.Value);
        if (context.LambdaDepth >= MaxCallDepth)
            return CellValue.Error(ErrorKind.Num);

        var scope = new Scope(closure.Captured);
        for (var i = 0; i < closure.Parameters.Length; i++)
            scope.Bind(closure.Parameters[i], Binding.Of(i < arguments.Count ? arguments[i] : CellValue.Missing));

        var saved = context.Scope;
        context.Scope = scope;
        context.LambdaDepth++;
        try
        {
            return Evaluator.Evaluate(closure.Body, context);
        }
        finally
        {
            context.LambdaDepth--;
            context.Scope = saved;
        }
    }

    /// <summary>Reads a LET name or LAMBDA parameter; a LET value is evaluated on first use.</summary>
    public static Operand Read(Binding binding, EvaluationContext context)
    {
        if (binding.Value is { } value)
            return value;
        if (binding.Evaluating)
            return CellValue.Error(ErrorKind.Name);

        var saved = context.Scope;
        context.Scope = binding.DefinitionScope;
        binding.Evaluating = true;
        try
        {
            var result = Evaluator.Evaluate(binding.Node!, context);
            binding.Value = result;
            return result;
        }
        finally
        {
            binding.Evaluating = false;
            context.Scope = saved;
        }
    }

    // LET(name1, value1, ..., calculation). Each value sees only the names before it, so every
    // binding gets its own scope in a chain.
    private static Operand Let(IReadOnlyList<FormulaNode> arguments, EvaluationContext context)
    {
        if (arguments.Count < 3 || arguments.Count % 2 == 0)
            return CellValue.Error(ErrorKind.Value);

        var scope = context.Scope;
        for (var i = 0; i < arguments.Count - 1; i += 2)
        {
            if (arguments[i] is not NameNode { Sheet: null } name)
                return CellValue.Error(ErrorKind.Value);

            var inner = new Scope(scope);
            inner.Bind(name.Name.ToUpperInvariant(), Binding.Lazy(arguments[i + 1], scope));
            scope = inner;
        }

        var saved = context.Scope;
        context.Scope = scope;
        try
        {
            return Evaluator.Evaluate(arguments[^1], context);
        }
        finally
        {
            context.Scope = saved;
        }
    }

    // LAMBDA(param..., body). An optional parameter is written [name].
    private static Operand Lambda(IReadOnlyList<FormulaNode> arguments, EvaluationContext context)
    {
        if (arguments.Count < 1)
            return CellValue.Error(ErrorKind.Value);

        var count = arguments.Count - 1;
        var names = new string[count];
        var optional = new bool[count];
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (var i = 0; i < count; i++)
        {
            string? name = arguments[i] switch
            {
                NameNode { Sheet: null } n => n.Name,
                StructuredReferenceNode s when IsOptionalParameter(s.Text, out var inner) => inner,
                _ => null,
            };
            if (name is null || !seen.Add(name.ToUpperInvariant()))
                return CellValue.Error(ErrorKind.Value);

            names[i] = name.ToUpperInvariant();
            optional[i] = arguments[i] is StructuredReferenceNode;
        }

        return CellValue.Lambda(new Closure(names, optional, arguments[^1], context.Scope));
    }

    // "[y]" lexes as a structured reference; files may write "[_xlpm.y]".
    private static bool IsOptionalParameter(string text, out string name)
    {
        name = "";
        if (text.Length < 3 || text[0] != '[' || text[^1] != ']')
            return false;

        var inner = text[1..^1];
        if (inner.StartsWith("_xlpm.", StringComparison.OrdinalIgnoreCase))
            inner = inner[6..];
        if (inner.Length == 0 || inner.AsSpan().IndexOfAny("[]'#@ ,") >= 0)
            return false;

        name = inner;
        return true;
    }
}
