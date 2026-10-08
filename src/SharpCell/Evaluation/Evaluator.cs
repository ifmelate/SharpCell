using System.Collections.Generic;
using SharpCell.Functions;
using SharpCell.Parsing;

namespace SharpCell.Evaluation;

/// <summary>
/// Evaluates formula trees. Recursion follows the tree, whose depth the parser bounds
/// (<see cref="FormulaLimits.MaxTreeDepth"/>); left spines of operator chains are folded in a loop.
/// </summary>
internal static class Evaluator
{
    // Turning a reference into an array allocates every cell; beyond this the result is #NUM!
    // rather than an out-of-memory crash. A full column (about one million cells) fits.
    internal const long MaxArrayCells = 1L << 24;

    /// <summary>Evaluates a whole formula and turns the result into what a cell holds.</summary>
    public static CellValue EvaluateFormula(FormulaNode root, EvaluationContext context)
    {
        var node = root;
        while (node is ParenthesesNode p)
            node = p.Inner;

        var result = node is BinaryNode binary ? EvaluateBinary(binary, context, isRoot: true) : Evaluate(node, context);
        var value = ToValue(result, context);
        return value.Kind is CellValueKind.Empty or CellValueKind.Missing ? CellValue.Number(0) : value;
    }

    public static Operand Evaluate(FormulaNode node, EvaluationContext context)
    {
        switch (node)
        {
            case NumberNode n:
                return CellValue.Number(n.Value);
            case TextNode t:
                return CellValue.Text(t.Value);
            case BooleanNode b:
                return CellValue.Boolean(b.Value);
            case ErrorNode e:
                return CellValue.Error(e.Error);
            case MissingNode:
                return CellValue.Missing;
            case ArrayNode a:
                return CellValue.Array(a.Values);
            case ReferenceNode r:
                return EvaluateReference(r, context);
            case RefErrorNode:
                return CellValue.Error(ErrorKind.Ref);
            case NameNode n:
                return EvaluateName(n, context);
            case ParenthesesNode p:
                return Evaluate(p.Inner, context);
            case UnaryNode { Operator: UnaryOperator.Plus } u:
                return Evaluate(u.Operand, context);
            case UnaryNode u:
                var operand = ToValue(Evaluate(u.Operand, context), context);
                return u.Operator == UnaryOperator.Negate
                    ? ArrayMath.Map(operand, v => Operators.Negate(v, context.Culture))
                    : ArrayMath.Map(operand, v => Operators.Percent(v, context.Culture));
            case BinaryNode b:
                return EvaluateBinary(b, context, isRoot: false);

            case FunctionNode f:
                return FunctionInvoker.Invoke(f, context);

            // Tables are not evaluated in v0.1; spill, @ and lambda calls arrive in stage 3.
            case StructuredReferenceNode:
                return CellValue.Error(ErrorKind.Name);
            default:
                return CellValue.Error(ErrorKind.Calc);
        }
    }

    /// <summary>Reads a reference: one cell gives its value, a single area gives an array.</summary>
    public static CellValue ToValue(Operand operand, EvaluationContext context)
    {
        if (operand.Reference is not { } reference)
            return operand.Value;
        if (!reference.IsSingleArea)
            return CellValue.Error(ErrorKind.Value);

        var (sheet, area) = reference.Areas[0];
        if (area.IsSingleCell)
            return context.ReadCell(sheet, area.FirstRow, area.FirstColumn);
        if (area.CellCount > MaxArrayCells)
            return CellValue.Error(ErrorKind.Num);

        var values = new CellValue[area.Rows, area.Columns];
        foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
            values[cell.Row - area.FirstRow, cell.Column - area.FirstColumn] = context.ReadCell(sheet, cell);
        return CellValue.Array(values);
    }

    private static Operand EvaluateBinary(BinaryNode node, EvaluationContext context, bool isRoot)
    {
        var spine = new List<BinaryNode>();
        FormulaNode current = node;
        while (current is BinaryNode b)
        {
            spine.Add(b);
            current = b.Left;
        }

        var accumulator = Evaluate(current, context);
        for (var i = spine.Count - 1; i >= 0; i--)
        {
            var b = spine[i];
            var right = Evaluate(b.Right, context);
            accumulator = Apply(b.Operator, accumulator, right, context, last: isRoot && i == 0);
        }

        return accumulator;
    }

    private static Operand Apply(BinaryOperator op, Operand left, Operand right, EvaluationContext context, bool last)
    {
        switch (op)
        {
            case BinaryOperator.Range:
                return Operators.Range(left, right);
            case BinaryOperator.Intersect:
                return Operators.Intersect(left, right);
            case BinaryOperator.Union:
                return Operators.Union(left, right);
        }

        var a = ToValue(left, context);
        var b = ToValue(right, context);
        var culture = context.Culture;
        return ArrayMath.Map(a, b, (x, y) => Operators.Binary(op, x, y, culture, last));
    }

    private static Operand EvaluateReference(ReferenceNode node, EvaluationContext context)
    {
        var area = Area.Resolve(node.Area, context.Origin);
        if (node.Sheet is null)
            return context.Sheet is { } current ? Operand.Of(new Reference(current, area)) : CellValue.Error(ErrorKind.Ref);

        var workbook = context.Workbook;
        if (!workbook.TryGetSheet(node.Sheet.First, out var first))
            return CellValue.Error(ErrorKind.Ref);
        if (node.Sheet.Last is null)
            return Operand.Of(new Reference(first!, area));

        if (!workbook.TryGetSheet(node.Sheet.Last, out var last))
            return CellValue.Error(ErrorKind.Ref);

        var sheets = workbook.Sheets;
        int from = IndexOf(sheets, first!), to = IndexOf(sheets, last!);
        if (from > to)
            (from, to) = (to, from);
        var areas = new List<SheetArea>();
        for (var i = from; i <= to; i++)
            areas.Add(new SheetArea(sheets[i], area));
        return Operand.Of(new Reference(areas));
    }

    private static int IndexOf(IReadOnlyList<Worksheet> sheets, Worksheet sheet)
    {
        for (var i = 0; i < sheets.Count; i++)
        {
            if (sheets[i] == sheet)
                return i;
        }

        return -1;
    }

    // Lookup order: the sheet's own names, then workbook names. Relative references inside a name
    // were parsed at A1 and move with the cell that uses the name.
    private static Operand EvaluateName(NameNode node, EvaluationContext context)
    {
        var workbook = context.Workbook;
        Worksheet? scope = context.Sheet;
        if (node.Sheet is not null && !workbook.TryGetSheet(node.Sheet.First, out scope))
            return CellValue.Error(ErrorKind.Ref);

        var upper = node.Name.ToUpperInvariant();
        if (!(scope is not null && workbook.Names.TryGet(upper, scope, out var definition))
            && !workbook.Names.TryGet(upper, null, out definition))
            return CellValue.Error(ErrorKind.Name);

        if (!context.TryEnterName(definition!))
            return CellValue.Error(ErrorKind.Ref);
        try
        {
            return Evaluate(definition!.Formula, context);
        }
        finally
        {
            context.LeaveName();
        }
    }
}
