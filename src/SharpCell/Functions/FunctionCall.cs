using System.Collections.Generic;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Functions;

/// <summary>Arguments of one call as a function body sees them.</summary>
internal sealed class FunctionCall
{
    private readonly IReadOnlyList<FormulaNode> _nodes;
    private readonly Operand?[] _arguments;
    private readonly int _pendingAtStart;

    public FunctionCall(EvaluationContext context, IReadOnlyList<FormulaNode> nodes)
    {
        Context = context;
        _nodes = nodes;
        _arguments = new Operand?[nodes.Count];
        _pendingAtStart = context.Pending.Count;
    }

    /// <summary>Whether this call has read a dirty cell; its result will be discarded.</summary>
    public bool MetPendingInput => Context.Pending.Count > _pendingAtStart;

    public EvaluationContext Context { get; }

    public int Count => _nodes.Count;

    /// <summary>
    /// The argument as evaluated; a lazy argument is evaluated on first access. Once the call has
    /// read a dirty cell, a lazy argument is not evaluated at all: the deciding value was a stand-in,
    /// and following the branch it picked could read cells the real evaluation never reads (and
    /// report a loop that does not exist).
    /// </summary>
    public Operand this[int index]
    {
        get
        {
            if (_arguments[index] is { } evaluated)
                return evaluated;
            if (MetPendingInput)
                return EvaluationContext.PendingPlaceholder;
            return (_arguments[index] = Evaluator.Evaluate(_nodes[index], Context)).Value;
        }
    }

    /// <summary>The argument read as a value (a reference becomes its value or an array).</summary>
    public CellValue Value(int index) => Evaluator.ToValue(this[index], Context);

    /// <summary>Whether the argument was left empty, as the middle one in <c>IF(A1,,B1)</c>.</summary>
    public bool IsMissing(int index) => _nodes[index] is MissingNode;

    internal void Set(int index, Operand value) => _arguments[index] = value;
}
