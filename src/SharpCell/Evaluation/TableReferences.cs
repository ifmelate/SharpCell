using System;
using SharpCell.Parsing;

namespace SharpCell.Evaluation;

/// <summary>Turns a structured reference into the cells of the table it covers now.</summary>
internal static class TableReferences
{
    public static Operand Resolve(StructuredReference reference, EvaluationContext context)
    {
        if (FindTable(reference.Table, context) is not { } table)
            return CellValue.Error(ErrorKind.Ref);
        if (!TryRows(table, reference.Rows, context.Origin.Row, out var firstRow, out var lastRow, out var error))
            return CellValue.Error(error);

        var firstColumn = table.Area.FirstColumn;
        var lastColumn = table.Area.LastColumn;
        if (reference.FirstColumn is { } from)
        {
            var a = table.ColumnIndex(from);
            var b = table.ColumnIndex(reference.LastColumn ?? from);
            if (a < 0 || b < 0)
                return CellValue.Error(ErrorKind.Ref);
            firstColumn = table.Area.FirstColumn + Math.Min(a, b);
            lastColumn = table.Area.FirstColumn + Math.Max(a, b);
        }

        return Operand.Of(new Reference(table.Worksheet, new Area(firstRow, firstColumn, lastRow, lastColumn)));
    }

    // A named table is a dependency even while it does not exist: adding it recalculates the formula.
    // Without a name it is the table holding the formula; adding a table invalidates the cells inside it.
    private static Table? FindTable(string? name, EvaluationContext context)
    {
        if (name is null)
            return context.Sheet is { } sheet ? context.Workbook.TableAt(sheet, context.Origin.Row, context.Origin.Column) : null;

        context.RecordName(name.ToUpperInvariant());
        return context.Workbook.TryGetTable(name, out var table) ? table : null;
    }

    private static bool TryRows(Table table, TableRows rows, int formulaRow, out int first, out int last, out ErrorKind error)
    {
        int top = table.Area.FirstRow, bottom = table.Area.LastRow;
        (first, last) = rows switch
        {
            TableRows.All => (top, bottom),
            TableRows.Headers => (top, top),
            TableRows.Totals => (bottom, bottom),
            TableRows.Headers | TableRows.Data => (top, table.LastDataRow),
            TableRows.Data | TableRows.Totals => (table.FirstDataRow, bottom),
            TableRows.ThisRow => (formulaRow, formulaRow),
            _ => (table.FirstDataRow, table.LastDataRow),
        };

        error = ErrorKind.Ref;
        if (rows != TableRows.All && (rows & TableRows.Headers) != 0 && !table.HasHeaderRow)
            return false;
        if (rows != TableRows.All && (rows & TableRows.Totals) != 0 && !table.HasTotalsRow)
            return false;

        // #This Row needs a data row on the formula's own row; the formula's column does not matter.
        error = ErrorKind.Value;
        return rows != TableRows.ThisRow || (formulaRow >= table.FirstDataRow && formulaRow <= table.LastDataRow);
    }
}
