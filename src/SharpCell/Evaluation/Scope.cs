using System;
using System.Collections.Generic;
using SharpCell.Parsing;

namespace SharpCell.Evaluation;

/// <summary>One LET binding or LAMBDA parameter. A LET value is evaluated on first use and kept.</summary>
internal sealed class Binding
{
    private Binding(Operand? value, FormulaNode? node, Scope? scope)
    {
        Value = value;
        Node = node;
        DefinitionScope = scope;
    }

    public Operand? Value { get; set; }

    public FormulaNode? Node { get; }

    /// <summary>Where a lazy value is evaluated: the scope before its own name exists.</summary>
    public Scope? DefinitionScope { get; }

    public bool Evaluating { get; set; }

    /// <summary>The value was computed from a dirty cell: every later read must signal that again.</summary>
    public bool MetPending { get; set; }

    public static Binding Lazy(FormulaNode node, Scope? scope) => new(null, node, scope);

    public static Binding Of(Operand value) => new(value, null, null);
}

/// <summary>A lexical scope of LET names and LAMBDA parameters; lookups walk outward.</summary>
internal sealed class Scope(Scope? parent)
{
    private readonly Dictionary<string, Binding> _bindings = new(StringComparer.Ordinal);

    public void Bind(string upperName, Binding binding) => _bindings[upperName] = binding;

    public bool TryFind(string upperName, out Binding? binding)
    {
        for (var scope = this; scope is not null; scope = scope.Parent)
        {
            if (scope._bindings.TryGetValue(upperName, out binding))
                return true;
        }

        binding = null;
        return false;
    }

    private Scope? Parent { get; } = parent;
}

/// <summary>The function value made by LAMBDA: parameters, body and the scope it was created in.</summary>
internal sealed class Closure(string[] parameters, bool[] optional, FormulaNode body, Scope? captured) : LambdaValue
{
    public string[] Parameters { get; } = parameters;

    public bool[] Optional { get; } = optional;

    public FormulaNode Body { get; } = body;

    public Scope? Captured { get; } = captured;

    public int RequiredCount
    {
        get
        {
            var count = 0;
            foreach (var optional in Optional)
            {
                if (!optional)
                    count++;
            }

            return count;
        }
    }
}
