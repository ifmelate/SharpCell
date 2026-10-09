using System;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Functions;

/// <summary>
/// Functions that compute references or describe them. OFFSET and INDIRECT are volatile; every
/// reference a function returns is recorded as a dependency of this evaluation, so the graph
/// follows wherever it points now.
/// </summary>
internal static class ReferenceFunctions
{
    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("OFFSET", 3, 5, [ArgumentKind.Any, ArgumentKind.Value], Offset) { IsVolatile = true });
        registry.Add(new FunctionInfo("INDIRECT", 1, 2, [ArgumentKind.Value], Indirect) { IsVolatile = true });

        // INDEX's array is an array parameter: INDEX(A1:A9*2,3) worked in old Excel without Ctrl+Shift+Enter.
        registry.Add(new FunctionInfo("INDEX", 2, 4, [ArgumentKind.ArrayContext, ArgumentKind.Value], Index));
        registry.Add(new FunctionInfo("ROW", 0, 1, [ArgumentKind.Any], call => Position(call, rows: true)));
        registry.Add(new FunctionInfo("COLUMN", 0, 1, [ArgumentKind.Any], call => Position(call, rows: false)));
        registry.Add(new FunctionInfo("AREAS", 1, 1, [ArgumentKind.Any], Areas));
        registry.Add(new FunctionInfo("ADDRESS", 2, 5, [ArgumentKind.Value], Address));
        registry.Add(new FunctionInfo("FORMULATEXT", 1, 1, [ArgumentKind.Any], FormulaText)
        {
            Status = FunctionStatus.KnownDeviation,
            Deviation = "Formulas read from xlsx files are shown in SharpCell's canonical form, without the spaces and line breaks the author typed.",
        });
        registry.Add(new FunctionInfo("TRIMRANGE", 1, 3, [ArgumentKind.Any, ArgumentKind.Value], TrimRange));
    }

    // INDEX(array, row, [column], [area]): the element, or with row or column 0 (or left empty) the
    // whole column or row. A reference gives a reference, so SUM(INDEX(A1:C3,0,2)) and
    // INDEX(...):B5 work. With one index, a single row or column is indexed along its length.
    // A negative index is #VALUE!, one outside the array (or a missing area) #REF!.
    private static Operand Index(FunctionCall call)
    {
        var row = call.Integer(1, 0);
        if (row.IsError)
            return row;
        var column = call.Integer(2, 0);
        if (column.IsError)
            return column;
        var areaNumber = call.Integer(3, 1);
        if (areaNumber.IsError)
            return areaNumber;

        var argument = call[0];
        LookupTable table;
        if (argument.Reference is { } reference)
        {
            if (areaNumber.AsNumber() < 1 || areaNumber.AsNumber() > reference.Areas.Count)
                return CellValue.Error(ErrorKind.Ref);
            var (sheet, area) = reference.Areas[(int)areaNumber.AsNumber() - 1];
            LookupTable.TryCreate(Operand.Of(new Reference(sheet, area)), out table, out _);
        }
        else
        {
            if (!LookupTable.TryCreate(argument, out table, out var error))
                return error;
            if (areaNumber.AsNumber() != 1)
                return CellValue.Error(ErrorKind.Ref);
        }

        double r = row.AsNumber(), c = column.AsNumber();
        if (r < 0 || c < 0)
            return CellValue.Error(ErrorKind.Value);
        if (call.Count < 3 && table.Rows == 1)
            (r, c) = (1, r);
        if (r > table.Rows || c > table.Columns)
            return CellValue.Error(ErrorKind.Ref);

        return table.Slice(
            r == 0 ? 0 : (int)r - 1, c == 0 ? 0 : (int)c - 1,
            r == 0 ? table.Rows : 1, c == 0 ? table.Columns : 1, call.Context);
    }

    // ROW([reference]) and COLUMN([reference]): of the formula's cell without an argument; a
    // reference of several rows (columns) gives a column (row) of their numbers.
    private static Operand Position(FunctionCall call, bool rows)
    {
        if (call.Count == 0)
            return CellValue.Number(rows ? call.Context.Origin.Row : call.Context.Origin.Column);

        var argument = call[0];
        if (argument.Reference is not { } reference)
            return argument.Value.IsError ? argument.Value : CellValue.Error(ErrorKind.Value);
        if (!reference.IsSingleArea)
            return CellValue.Error(ErrorKind.Ref);

        var area = reference.Areas[0].Area;
        var first = rows ? area.FirstRow : area.FirstColumn;
        var count = rows ? area.Rows : area.Columns;
        if (count == 1)
            return CellValue.Number(first);

        var values = rows ? new CellValue[count, 1] : new CellValue[1, count];
        for (var i = 0; i < count; i++)
        {
            if (rows)
                values[i, 0] = CellValue.Number(first + i);
            else
                values[0, i] = CellValue.Number(first + i);
        }

        return CellValue.Array(values);
    }

    private static Operand Areas(FunctionCall call)
    {
        var argument = call[0];
        if (argument.Reference is { } reference)
            return CellValue.Number(reference.Areas.Count);
        return argument.Value.IsError ? argument.Value : CellValue.Error(ErrorKind.Value);
    }

    // ADDRESS(row, column, [abs], [a1], [sheet]): abs 1 is $A$1, 2 A$1, 3 $A1, 4 A1; R1C1 style
    // writes relative parts in brackets. The sheet name is quoted when it has to be.
    private static Operand Address(FunctionCall call)
    {
        var row = call.Integer(0);
        if (row.IsError)
            return row;
        var column = call.Integer(1);
        if (column.IsError)
            return column;
        var kind = call.Integer(2, 1);
        if (kind.IsError)
            return kind;
        var a1 = call.Boolean(3, true);
        if (a1.IsError)
            return a1;
        var sheet = call.Has(4) ? call.Text(4) : CellValue.Empty;
        if (sheet.IsError)
            return sheet;

        double r = row.AsNumber(), c = column.AsNumber(), k = kind.AsNumber();
        if (k is < 1 or > 4 || r < 1 || r > CellAddress.MaxRow || c < 1 || c > CellAddress.MaxColumn)
            return CellValue.Error(ErrorKind.Value);

        var absoluteRow = k is 1 or 2;
        var absoluteColumn = k is 1 or 3;
        var rowText = ((int)r).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var columnNumber = (int)c;
        string address;
        if (a1.AsBoolean())
        {
            address = (absoluteColumn ? "$" : "") + ColumnNames.ToName(columnNumber) + (absoluteRow ? "$" : "") + rowText;
        }
        else
        {
            var columnText = columnNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
            address = "R" + (absoluteRow ? rowText : "[" + rowText + "]") + "C" + (absoluteColumn ? columnText : "[" + columnText + "]");
        }

        if (sheet.Kind != CellValueKind.Text)
            return CellValue.Text(address);
        var name = sheet.AsText();
        return CellValue.Text((SheetNeedsQuotes(name) ? "'" + name.Replace("'", "''", StringComparison.Ordinal) + "'" : name) + "!" + address);
    }

    // Unquoted sheet names must read as one word and not as a reference or a boolean.
    private static bool SheetNeedsQuotes(string name)
    {
        if (name.Length == 0)
            return false;
        if (!(char.IsLetter(name[0]) || name[0] == '_'))
            return true;
        foreach (var ch in name)
        {
            if (!(char.IsLetterOrDigit(ch) || ch is '_' or '.'))
                return true;
        }

        return CellAddress.TryParse(name, out _)
            || ReferenceSyntax.TryParseR1C1Area(name, out _)
            || name.Equals("TRUE", StringComparison.OrdinalIgnoreCase)
            || name.Equals("FALSE", StringComparison.OrdinalIgnoreCase);
    }

    // FORMULATEXT(reference): the formula of the top-left cell, #N/A for a cell without one. It
    // reads the formula, not the value, so a formula can show its own text without a loop.
    private static Operand FormulaText(FunctionCall call)
    {
        var argument = call[0];
        if (argument.Reference is not { } reference)
            return argument.Value.IsError ? argument.Value : CellValue.Error(ErrorKind.Value);

        var (sheet, area) = reference.Areas[0];
        var cell = new Reference(sheet, Area.Cell(area.FirstRow, area.FirstColumn));
        call.Context.RecordReference(cell);
        return sheet.Store.Get(area.FirstRow, area.FirstColumn)?.FormulaText is { } text
            ? CellValue.Text(text)
            : CellValue.Error(ErrorKind.NA);
    }

    // TRIMRANGE(range, [rows], [columns]): the range without its empty outer rows and columns;
    // 0 trims nothing, 1 leading, 2 trailing, 3 (default) both. Only stored cells are visited.
    private static Operand TrimRange(FunctionCall call)
    {
        var argument = call[0];
        if (argument.Reference is not { } reference)
            return argument.Value.IsError ? argument.Value : CellValue.Error(ErrorKind.Value);
        if (!reference.IsSingleArea)
            return CellValue.Error(ErrorKind.Value);
        var trimRows = call.Integer(1, 3);
        if (trimRows.IsError)
            return trimRows;
        var trimColumns = call.Integer(2, 3);
        if (trimColumns.IsError)
            return trimColumns;
        if (trimRows.AsNumber() is < 0 or > 3 || trimColumns.AsNumber() is < 0 or > 3)
            return CellValue.Error(ErrorKind.Value);

        var context = call.Context;
        var (sheet, area) = reference.Areas[0];
        int firstRow = int.MaxValue, firstColumn = int.MaxValue, lastRow = 0, lastColumn = 0, visited = 0;
        foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
        {
            if (++visited % 4096 == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            if (context.ReadCell(sheet, cell).Kind == CellValueKind.Empty)
                continue;
            firstRow = Math.Min(firstRow, cell.Row);
            lastRow = Math.Max(lastRow, cell.Row);
            firstColumn = Math.Min(firstColumn, cell.Column);
            lastColumn = Math.Max(lastColumn, cell.Column);
        }

        // Nothing to keep: Excel gives the range's first cell.
        if (lastRow == 0)
            return Trimmed(context, sheet, Area.Cell(area.FirstRow, area.FirstColumn));

        var rowMode = (int)trimRows.AsNumber();
        var columnMode = (int)trimColumns.AsNumber();
        return Trimmed(context, sheet, new Area(
            rowMode is 1 or 3 ? firstRow : area.FirstRow,
            columnMode is 1 or 3 ? firstColumn : area.FirstColumn,
            rowMode is 2 or 3 ? lastRow : area.LastRow,
            columnMode is 2 or 3 ? lastColumn : area.LastColumn));
    }

    private static Operand Trimmed(EvaluationContext context, Worksheet sheet, Area area)
    {
        var result = new Reference(sheet, area);
        context.RecordReference(result);
        return Operand.Of(result);
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

    // A range joining a cell to whole rows or columns, as R1C1:R2, is not a reference to INDIRECT.
    private static bool DenotesReference(FormulaNode node) => node switch
    {
        ReferenceNode or NameNode or RefErrorNode => true,
        BinaryNode { Operator: BinaryOperator.Range, Left: ReferenceNode left, Right: ReferenceNode right } =>
            Family(left.Area.Kind) == Family(right.Area.Kind),
        BinaryNode { Operator: BinaryOperator.Range } range => DenotesReference(range.Left) && DenotesReference(range.Right),
        _ => false,
    };

    private static int Family(AreaKind kind) => kind is AreaKind.Cell or AreaKind.Range ? 0 : (int)kind;

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
