using System.Collections.Generic;
using System.Runtime.CompilerServices;
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
        if (context.Legacy)
            result = ImplicitIntersection(result, context);
        var value = ToValue(result, context);
        return value.Kind switch
        {
            CellValueKind.Empty or CellValueKind.Missing => CellValue.Number(0),

            // A function value cannot be shown in a cell.
            CellValueKind.Lambda => CellValue.Error(ErrorKind.Calc),
            _ => value,
        };
    }

    public static Operand Evaluate(FormulaNode node, EvaluationContext context)
    {
        // Tree depth is bounded per formula, but names and functions nest formulas inside formulas.
        // Rather than overflow the thread stack, a formula that nests too deep evaluates to #NUM!.
        if (!RuntimeHelpers.TryEnsureSufficientExecutionStack())
            return CellValue.Error(ErrorKind.Num);

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
                var operand = ToValue(LegacyScalar(Evaluate(u.Operand, context), context), context);
                return u.Operator == UnaryOperator.Negate
                    ? ArrayMath.Map(operand, v => Operators.Negate(v, context.Culture, context.DateSystem))
                    : ArrayMath.Map(operand, v => Operators.Percent(v, context.Culture, context.DateSystem));
            case BinaryNode b:
                return EvaluateBinary(b, context, isRoot: false);

            case FunctionNode f:
                return EvaluateCall(f, context);
            case CallNode c:
                return Lambdas.Call(ToValue(Evaluate(c.Callee, context), context), c.Arguments, context);

            case SpillNode s:
                return EvaluateSpill(s, context);
            case ImplicitIntersectionNode i:
                return ImplicitIntersection(Evaluate(i.Operand, context), context);

            case StructuredReferenceNode s:
                return EvaluateStructured(s.Reference, context);
            case UnsupportedNode u:
                context.Report(DiagnosticKind.UnsupportedFormula, u.Reason);
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
                // The bounding box covers cells neither side named (B2 in A1:C3), so it is recorded too.
                var range = Operators.Range(left, right);
                if (range.Reference is { } box)
                    context.RecordReference(box);
                return range;
            case BinaryOperator.Intersect:
                return Operators.Intersect(left, right);
            case BinaryOperator.Union:
                return Operators.Union(left, right);
        }

        var a = ToValue(LegacyScalar(left, context), context);
        var b = ToValue(LegacyScalar(right, context), context);
        var culture = context.Culture;
        var dateSystem = context.DateSystem;
        return ArrayMath.Map(a, b, (x, y) => Operators.Binary(op, x, y, culture, dateSystem, last));
    }

    /// <summary>
    /// The @ operator. A range gives the cell in the formula's row (single column), column (single
    /// row) or both; no such cell is #VALUE!. An array gives its top-left element.
    /// </summary>
    /// <summary>
    /// In a formula from before dynamic arrays, a range where one value is expected (an operator's
    /// operand, a scalar function argument) is reduced to the cell in the formula's row or column,
    /// which Excel shows as <c>@A1:A3</c>. Array constants and array results stay arrays.
    /// </summary>
    public static Operand LegacyScalar(Operand operand, EvaluationContext context) =>
        context.Legacy && operand.Reference is { } reference && !(reference.IsSingleArea && reference.Areas[0].Area.IsSingleCell)
            ? ImplicitIntersection(operand, context)
            : operand;

    public static Operand ImplicitIntersection(Operand operand, EvaluationContext context)
    {
        if (operand.Reference is not { } reference)
        {
            var value = operand.Value;
            return value.Kind == CellValueKind.Array ? value.AsArray()[0, 0] : value;
        }

        if (!reference.IsSingleArea)
            return CellValue.Error(ErrorKind.Value);

        var (sheet, area) = reference.Areas[0];
        if (area.IsSingleCell)
            return operand;

        var row = area.Rows == 1 ? area.FirstRow : context.Origin.Row;
        var column = area.Columns == 1 ? area.FirstColumn : context.Origin.Column;
        if (!area.Contains(row, column))
            return CellValue.Error(ErrorKind.Value);

        var cell = new Reference(sheet, Area.Cell(row, column));
        context.RecordReference(cell);
        return Operand.Of(cell);
    }

    // A1#: the area the array result of anchor A1 currently covers.
    private static Operand EvaluateSpill(SpillNode node, EvaluationContext context)
    {
        var operand = Evaluate(node.Operand, context);
        if (operand.Reference is not { IsSingleArea: true } reference || !reference.Areas[0].Area.IsSingleCell)
            return operand.IsReference || !operand.Value.IsError ? CellValue.Error(ErrorKind.Ref) : operand;

        var (sheet, area) = reference.Areas[0];
        var anchor = sheet.Store.Get(area.FirstRow, area.FirstColumn);
        if (anchor is { IsDirty: true, Formula: not null })
            return context.ReadCell(sheet, area.FirstRow, area.FirstColumn);
        // An array formula (Ctrl+Shift+Enter) counts too: Excel gives A1# its whole fixed area.
        if (anchor?.SpillArea is not { } spill)
            return CellValue.Error(ErrorKind.Ref);

        var result = new Reference(sheet, spill);
        context.RecordReference(result);
        return Operand.Of(result);
    }

    private static Operand EvaluateReference(ReferenceNode node, EvaluationContext context)
    {
        var result = ResolveReference(node, context);
        if (result.Reference is { } reference)
            context.RecordReference(reference);
        return result;
    }

    private static Operand EvaluateStructured(StructuredReference reference, EvaluationContext context)
    {
        var result = TableReferences.Resolve(reference, context);
        if (result.Reference is { } resolved)
            context.RecordReference(resolved);
        return result;
    }

    private static Operand ResolveReference(ReferenceNode node, EvaluationContext context)
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

    // A name followed by arguments: LET/LAMBDA syntax, a LET-bound function, a registry function,
    // or a defined name holding a LAMBDA, in that order.
    private static Operand EvaluateCall(FunctionNode node, EvaluationContext context)
    {
        if (Lambdas.TryEvaluateSpecialForm(node, context, out var special))
            return special;

        if (context.Scope is { } scope && scope.TryFind(node.Name, out var binding))
            return Lambdas.Call(ToValue(Lambdas.Read(binding!, context), context), node.Arguments, context);

        if (context.Workbook.Functions.TryGet(node.Name, out _))
            return FunctionInvoker.Invoke(node, context);

        var name = EvaluateName(new NameNode(null, node.Name), context);
        if (!name.IsReference && name.Value.Kind == CellValueKind.Error && name.Value.AsError() == ErrorKind.Name)
            return name;
        return Lambdas.Call(ToValue(name, context), node.Arguments, context);
    }

    // Lookup order: LET names and LAMBDA parameters, the sheet's own names, then workbook names. Relative references inside a name
    // were parsed at A1 and move with the cell that uses the name.
    private static Operand EvaluateName(NameNode node, EvaluationContext context)
    {
        var workbook = context.Workbook;
        Worksheet? scope = context.Sheet;
        if (node.Sheet is not null && !workbook.TryGetSheet(node.Sheet.First, out scope))
            return CellValue.Error(ErrorKind.Ref);

        var upper = node.Name.ToUpperInvariant();
        if (node.Sheet is null && context.Scope is { } local && local.TryFind(upper, out var binding))
            return Lambdas.Read(binding!, context);

        context.RecordName(upper);
        if (!(scope is not null && workbook.Names.TryGet(upper, scope, out var definition))
            && !workbook.Names.TryGet(upper, null, out definition))
        {
            // A table name alone is its data rows: SUM(Sales) is SUM(Sales[]).
            if (node.Sheet is null && workbook.TryGetTable(upper, out _))
                return EvaluateStructured(new StructuredReference(node.Name, TableRows.Data, null, null), context);
            return CellValue.Error(ErrorKind.Name);
        }

        if (!context.TryEnterName(definition!))
            return CellValue.Error(ErrorKind.Ref);

        // A defined name is evaluated on its own terms: the caller's LET names and LAMBDA
        // parameters are not visible to it (lambdas it creates capture no caller scope).
        var callerScope = context.Scope;
        context.Scope = null;
        try
        {
            return Evaluate(definition!.Formula, context);
        }
        finally
        {
            context.Scope = callerScope;
            context.LeaveName();
        }
    }
}
