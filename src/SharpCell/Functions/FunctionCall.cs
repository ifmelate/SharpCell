using System;
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
        _pendingAtStart = context.PendingReads;
    }

    /// <summary>Whether this call has read a dirty cell; its result will be discarded.</summary>
    public bool MetPendingInput => Context.PendingReads > _pendingAtStart;

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

    /// <summary>Whether the call has the argument and it was not left empty.</summary>
    public bool Has(int index) => index < Count && !IsMissing(index);

    // Typed readers: each returns the converted value or the error to return, so a body reads
    //     var x = call.Number(0); if (x.IsError) return x;
    // An argument left empty in the call reads as 0, "" or FALSE; the overloads with a default use
    // the default both when the argument is absent and when it is left empty.

    /// <summary>The argument as a number by Excel's coercion (text such as "2" or "12:00" included).</summary>
    public CellValue Number(int index) => Coercion.ToNumber(Value(index), Context.Culture, Context.DateSystem);

    public CellValue Number(int index, double absent) => Has(index) ? Number(index) : CellValue.Number(absent);

    /// <summary>The argument as a number truncated toward zero, as Excel reads counts and positions.</summary>
    public CellValue Integer(int index)
    {
        var number = Number(index);
        return number.IsError ? number : CellValue.Number(Math.Truncate(number.AsNumber()));
    }

    public CellValue Integer(int index, double absent) => Has(index) ? Integer(index) : CellValue.Number(absent);

    public CellValue Text(int index) => Coercion.ToText(Value(index), Context.Culture);

    public CellValue Text(int index, string absent) => Has(index) ? Text(index) : CellValue.Text(absent);

    public CellValue Boolean(int index) => Coercion.ToBoolean(Value(index));

    public CellValue Boolean(int index, bool absent) => Has(index) ? Boolean(index) : CellValue.Boolean(absent);

    internal void Set(int index, Operand value) => _arguments[index] = value;
}
