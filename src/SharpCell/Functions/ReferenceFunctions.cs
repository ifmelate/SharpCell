using System;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Functions;

/// <summary>
/// Functions that compute references. Both are volatile, and the reference they return is recorded
/// as a dependency of this evaluation, so the graph follows wherever they point now.
/// </summary>
internal static class ReferenceFunctions
{
    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("OFFSET", 3, 5, [ArgumentKind.Any, ArgumentKind.Value], Offset) { IsVolatile = true });
        registry.Add(new FunctionInfo("INDIRECT", 1, 2, [ArgumentKind.Value], Indirect) { IsVolatile = true });
    }

    // OFFSET(reference, rows, columns, [height], [width]): height and width default to the
    // reference's size. A negative size extends up or left from the moved corner; a size between 0
    // and 1 counts as 1. A size of 0 or a result outside the sheet is #REF!.
    private static Operand Offset(FunctionCall call)
    {
        if (call[0].Reference is not { } reference)
            return call[0].Value.IsError ? call[0].Value : CellValue.Error(ErrorKind.Value);
        if (!reference.IsSingleArea)
            return CellValue.Error(ErrorKind.Value);

        var (sheet, area) = reference.Areas[0];
        if (!TryGetInteger(call, 1, 0, out var rows, out var error)
            || !TryGetInteger(call, 2, 0, out var columns, out error)
            || !TryGetInteger(call, 3, area.Rows, out var height, out error, isSize: true)
            || !TryGetInteger(call, 4, area.Columns, out var width, out error, isSize: true))
            return error;
        if (height == 0 || width == 0)
            return CellValue.Error(ErrorKind.Ref);

        var row = area.FirstRow + rows;
        var column = area.FirstColumn + columns;
        var firstRow = height > 0 ? row : row + height + 1;
        var firstColumn = width > 0 ? column : column + width + 1;
        var lastRow = firstRow + Math.Abs(height) - 1;
        var lastColumn = firstColumn + Math.Abs(width) - 1;
        if (firstRow < 1 || firstColumn < 1 || lastRow > CellAddress.MaxRow || lastColumn > CellAddress.MaxColumn)
            return CellValue.Error(ErrorKind.Ref);

        var result = new Reference(sheet, new Area((int)firstRow, (int)firstColumn, (int)lastRow, (int)lastColumn));
        call.Context.RecordReference(result);
        return Operand.Of(result);
    }

    // INDIRECT(text, [a1]): text naming a reference in A1 (default) or R1C1 style, relative to the
    // formula's cell. Defined names work; LET names do not (they are not visible to text).
    private static Operand Indirect(FunctionCall call)
    {
        var context = call.Context;
        var text = Coercion.ToText(call.Value(0), context.Culture);
        if (text.IsError)
            return text;

        // An empty style argument, as in INDIRECT("R2C2",), is FALSE: R1C1.
        var style = ReferenceStyle.A1;
        if (call.Count > 1 && call.IsMissing(1))
        {
            style = ReferenceStyle.R1C1;
        }
        else if (call.Count > 1)
        {
            var a1 = Coercion.ToBoolean(call.Value(1));
            if (a1.IsError)
                return a1;
            style = a1.AsBoolean() ? ReferenceStyle.A1 : ReferenceStyle.R1C1;
        }

        var source = text.AsText();
        if (source.Length == 0 || source.StartsWith('='))
            return CellValue.Error(ErrorKind.Ref);

        FormulaNode node;
        try
        {
            node = FormulaParser.Parse(source, context.Origin, style);
        }
        catch (FormulaParseException)
        {
            return CellValue.Error(ErrorKind.Ref);
        }

        if (!DenotesReference(node))
            return CellValue.Error(ErrorKind.Ref);

        var saved = context.Scope;
        context.Scope = null;
        try
        {
            var result = Evaluator.Evaluate(node, context);
            return result.IsReference ? result : CellValue.Error(ErrorKind.Ref);
        }
        finally
        {
            context.Scope = saved;
        }
    }

    private static bool DenotesReference(FormulaNode node) => node switch
    {
        ReferenceNode or NameNode or RefErrorNode => true,
        BinaryNode { Operator: BinaryOperator.Range } range => DenotesReference(range.Left) && DenotesReference(range.Right),
        _ => false,
    };

    private static bool TryGetInteger(FunctionCall call, int index, long absent, out long value, out CellValue error, bool isSize = false)
    {
        error = default;
        value = absent;
        if (index >= call.Count || call.IsMissing(index))
            return true;

        var number = Coercion.ToNumber(call.Value(index), call.Context.Culture, call.Context.DateSystem);
        if (number.IsError)
        {
            error = number;
            return false;
        }

        var truncated = Math.Truncate(number.AsNumber());
        if (isSize && truncated == 0 && number.AsNumber() != 0)
            truncated = Math.Sign(number.AsNumber());
        if (Math.Abs(truncated) > CellAddress.MaxRow * 2.0)
        {
            error = CellValue.Error(ErrorKind.Ref);
            return false;
        }

        value = (long)truncated;
        return true;
    }
}
