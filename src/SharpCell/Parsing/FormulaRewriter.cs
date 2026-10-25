using System.Collections.Generic;

namespace SharpCell.Parsing;

/// <summary>
/// Walks a formula and rebuilds the parts a subclass replaces; nodes nothing in changed are kept,
/// so an unchanged formula comes back as the same instance.
/// </summary>
internal abstract class FormulaRewriter
{
    /// <summary>The node to put in place of <paramref name="node"/>, or null to walk into it.</summary>
    protected abstract FormulaNode? Replace(FormulaNode node);

    public FormulaNode Visit(FormulaNode node)
    {
        if (Replace(node) is { } replaced)
            return replaced;
        switch (node)
        {
            case UnaryNode unary:
                return Visit(unary.Operand) is var operand && operand != unary.Operand ? new UnaryNode(unary.Operator, operand) : node;
            case BinaryNode binary:
                return Binary(binary);
            case ParenthesesNode parentheses:
                return Visit(parentheses.Inner) is var inner && inner != parentheses.Inner ? new ParenthesesNode(inner) : node;
            case FunctionNode function:
                return Arguments(function.Arguments) is { } changed ? new FunctionNode(function.Name, changed, function.FilePrefix) : node;
            case CallNode call:
            {
                var callee = Visit(call.Callee);
                var arguments = Arguments(call.Arguments);
                return callee != call.Callee || arguments is not null ? new CallNode(callee, arguments ?? call.Arguments) : node;
            }
            case SpillNode spill:
                return Visit(spill.Operand) is var spilled && spilled != spill.Operand ? new SpillNode(spilled) : node;
            case ImplicitIntersectionNode intersection:
                return Visit(intersection.Operand) is var operandNode && operandNode != intersection.Operand ? new ImplicitIntersectionNode(operandNode) : node;
            default:
                return node;
        }
    }

    // Operator chains are walked along their left spine in a loop: a long A1+A2+...+An does not
    // recurse once per operator.
    private FormulaNode Binary(BinaryNode node)
    {
        var chain = new List<BinaryNode> { node };
        var current = node;
        while (current.Left is BinaryNode left && Replace(left) is null)
        {
            chain.Add(left);
            current = left;
        }

        var result = Visit(current.Left);
        for (var i = chain.Count - 1; i >= 0; i--)
        {
            var binary = chain[i];
            var right = Visit(binary.Right);
            result = result == binary.Left && right == binary.Right ? binary : new BinaryNode(binary.Operator, result, right);
        }

        return result;
    }

    private List<FormulaNode>? Arguments(IReadOnlyList<FormulaNode> arguments)
    {
        List<FormulaNode>? changed = null;
        for (var i = 0; i < arguments.Count; i++)
        {
            var argument = Visit(arguments[i]);
            if (argument != arguments[i] && changed is null)
            {
                changed = new List<FormulaNode>(arguments.Count);
                for (var j = 0; j < i; j++)
                    changed.Add(arguments[j]);
            }

            changed?.Add(argument);
        }

        return changed;
    }
}
