using System.Collections.Generic;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Functions;

/// <summary>Arguments of one call as a function body sees them.</summary>
internal sealed class FunctionCall
{
    private readonly IReadOnlyList<FormulaNode> _nodes;
    private readonly Operand?[] _arguments;

    public FunctionCall(EvaluationContext context, IReadOnlyList<FormulaNode> nodes)
    {
        Context = context;
        _nodes = nodes;
        _arguments = new Operand?[nodes.Count];
    }

    public EvaluationContext Context { get; }

    public int Count => _nodes.Count;

    /// <summary>The argument as evaluated; a lazy argument is evaluated on first access.</summary>
    public Operand this[int index] => _arguments[index] ??= Evaluator.Evaluate(_nodes[index], Context);

    /// <summary>The argument read as a value (a reference becomes its value or an array).</summary>
    public CellValue Value(int index) => Evaluator.ToValue(this[index], Context);

    /// <summary>Whether the argument was left empty, as the middle one in <c>IF(A1,,B1)</c>.</summary>
    public bool IsMissing(int index) => _nodes[index] is MissingNode;

    internal void Set(int index, Operand value) => _arguments[index] = value;
}
