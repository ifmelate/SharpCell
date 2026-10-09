using System;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>MATCH, VLOOKUP, HLOOKUP, LOOKUP, XLOOKUP, XMATCH. The searching itself is in <see cref="LookupSearch"/>.</summary>
internal static class LookupFunctions
{
    private const string RegexDeviation = "Match mode 3 (regular expressions) uses .NET regular expressions, whose syntax differs from Excel's PCRE2 in details.";

    public static void Register(FunctionRegistry registry)
    {
        // Lookup values are scalars, so an array of them gives an array of results. LOOKUP's vectors
        // are array parameters (LOOKUP(2,1/(A1:A9<>""),A1:A9) worked in old Excel without
        // Ctrl+Shift+Enter); MATCH's array is not.
        registry.Add(new FunctionInfo("MATCH", 2, 3, [ArgumentKind.Value, ArgumentKind.Any, ArgumentKind.Value], Match));
        registry.Add(new FunctionInfo("VLOOKUP", 3, 4, [ArgumentKind.Value, ArgumentKind.Any, ArgumentKind.Value], call => TableLookup(call, vertical: true)));
        registry.Add(new FunctionInfo("HLOOKUP", 3, 4, [ArgumentKind.Value, ArgumentKind.Any, ArgumentKind.Value], call => TableLookup(call, vertical: false)));
        registry.Add(new FunctionInfo("LOOKUP", 2, 3, [ArgumentKind.Value, ArgumentKind.ArrayContext], Lookup));
        registry.Add(new FunctionInfo("XLOOKUP", 3, 6, [ArgumentKind.Value, ArgumentKind.Any, ArgumentKind.Any, ArgumentKind.Any, ArgumentKind.Value], XLookup)
        {
            Status = FunctionStatus.KnownDeviation,
            Deviation = RegexDeviation,
        });
        registry.Add(new FunctionInfo("XMATCH", 2, 4, [ArgumentKind.Value, ArgumentKind.Any, ArgumentKind.Value], XMatch)
        {
            Status = FunctionStatus.KnownDeviation,
            Deviation = RegexDeviation,
        });
    }

    // MATCH(value, array, [type]): type 0 is exact (with wildcards), 1 (default) the largest value
    // not above it in ascending data, -1 the smallest value not below it in descending data.
    private static Operand Match(FunctionCall call)
    {
        var target = call.Value(0);
        if (target.IsError)
            return target;
        if (!LookupTable.TryCreate(call[1], out var table, out var error))
            return error;
        var type = call.Number(2, 1);
        if (type.IsError)
            return type;
        if (!table.IsVector)
            return CellValue.Error(ErrorKind.NA);

        target = LookupSearch.LegacyTarget(target);
        var vector = table.AsVector();
        var index = type.AsNumber() == 0
            ? LookupSearch.FindExact(vector, target, call.Context)
            : LookupSearch.FindSorted(vector, target, descending: type.AsNumber() < 0, call.Context);
        return index < 0 ? CellValue.Error(ErrorKind.NA) : CellValue.Number(index + 1);
    }

    // VLOOKUP(value, table, column, [approximate]) searches the first column and returns from the
    // given one; HLOOKUP the same with rows. A column outside the table is #REF!, below 1 #VALUE!.
    private static Operand TableLookup(FunctionCall call, bool vertical)
    {
        var target = call.Value(0);
        if (target.IsError)
            return target;
        if (call[1] is { IsReference: false, Value.Kind: CellValueKind.Text })
            return CellValue.Error(ErrorKind.Value);
        if (!LookupTable.TryCreate(call[1], out var table, out var error))
            return error;
        var number = call.Integer(2);
        if (number.IsError)
            return number;
        var approximate = call.Boolean(3, true);
        if (approximate.IsError)
            return approximate;

        var offset = number.AsNumber();
        if (offset < 1)
            return CellValue.Error(ErrorKind.Value);
        if (offset > (vertical ? table.Columns : table.Rows))
            return CellValue.Error(ErrorKind.Ref);

        target = LookupSearch.LegacyTarget(target);
        var vector = vertical ? table.Column(0) : table.Row(0);
        var index = approximate.AsBoolean()
            ? LookupSearch.FindSorted(vector, target, descending: false, call.Context)
            : LookupSearch.FindExact(vector, target, call.Context);
        if (index < 0)
            return CellValue.Error(ErrorKind.NA);
        return vertical ? table.Get(index, (int)offset - 1, call.Context) : table.Get((int)offset - 1, index, call.Context);
    }

    // LOOKUP(value, vector, [result]) or LOOKUP(value, array): always the approximate binary search.
    // A two-dimensional array is searched along its longer side (the first row if it is wider than
    // tall, else the first column) and the result comes from the last row or column.
    private static Operand Lookup(FunctionCall call)
    {
        var target = call.Value(0);
        if (target.IsError)
            return target;
        if (!LookupTable.TryCreate(call[1], out var table, out var error))
            return error;

        var byRow = table.Columns > table.Rows;
        var vector = byRow ? table.Row(0) : table.Column(0);
        var index = LookupSearch.FindSorted(vector, LookupSearch.LegacyTarget(target), descending: false, call.Context);
        if (call.Count < 3)
        {
            if (index < 0)
                return CellValue.Error(ErrorKind.NA);
            return byRow ? table.Get(table.Rows - 1, index, call.Context) : table.Get(index, table.Columns - 1, call.Context);
        }

        if (!LookupTable.TryCreate(call[2], out var result, out error))
            return error;
        if (!result.IsVector)
            return CellValue.Error(ErrorKind.NA);
        if (index < 0)
            return CellValue.Error(ErrorKind.NA);

        var resultVector = result.AsVector();
        if (index < resultVector.Length)
            return resultVector.Get(index, call.Context);

        // A result range shorter than the lookup vector is read on past its end, as Excel does.
        if (result.Sheet is not { } sheet)
            return CellValue.Error(ErrorKind.NA);
        var row = (long)result.Area.FirstRow + (result.Rows == 1 && result.Columns > 1 ? 0 : index);
        var column = (long)result.Area.FirstColumn + (result.Rows == 1 && result.Columns > 1 ? index : 0);
        if (row > CellAddress.MaxRow || column > CellAddress.MaxColumn)
            return CellValue.Error(ErrorKind.NA);
        var cell = new Reference(sheet, Area.Cell((int)row, (int)column));
        call.Context.RecordReference(cell);
        return Operand.Of(cell);
    }

    // XLOOKUP(value, lookup, return, [if_not_found], [match_mode], [search_mode]). The return array
    // has the lookup array's length along it; the result is the matching row or column of it, a
    // reference when the return array is one (so XLOOKUP(...):XLOOKUP(...) is a range).
    private static Operand XLookup(FunctionCall call)
    {
        var target = call.Value(0);
        if (target.IsError)
            return target;
        if (!LookupTable.TryCreate(call[1], out var lookup, out var error))
            return error;
        if (!LookupTable.TryCreate(call[2], out var result, out error))
            return error;
        if (!TryReadModes(call, 4, out var matchMode, out var searchMode, out error))
            return error;
        if (!lookup.IsVector)
            return CellValue.Error(ErrorKind.Value);

        // A single-cell lookup array goes along whichever side of the return array has length 1.
        var vertical = lookup.Rows > 1 || (lookup.Columns == 1 && result.Rows == 1);
        if (vertical ? result.Rows != lookup.Rows : result.Columns != lookup.Columns)
            return CellValue.Error(ErrorKind.Value);

        var found = LookupSearch.Find(lookup.AsVector(), target, matchMode, searchMode, call.Context);
        if (found is not { } index)
            return CellValue.Error(ErrorKind.Value);
        if (index < 0)
            return call.Has(3) ? call[3] : CellValue.Error(ErrorKind.NA);
        return vertical
            ? result.Slice(index, 0, 1, result.Columns, call.Context)
            : result.Slice(0, index, result.Rows, 1, call.Context);
    }

    // XMATCH(value, array, [match_mode], [search_mode]): the 1-based position.
    private static Operand XMatch(FunctionCall call)
    {
        var target = call.Value(0);
        if (target.IsError)
            return target;
        if (!LookupTable.TryCreate(call[1], out var lookup, out var error))
            return error;
        if (!TryReadModes(call, 2, out var matchMode, out var searchMode, out error))
            return error;
        if (!lookup.IsVector)
            return CellValue.Error(ErrorKind.Value);

        var found = LookupSearch.Find(lookup.AsVector(), target, matchMode, searchMode, call.Context);
        if (found is not { } index)
            return CellValue.Error(ErrorKind.Value);
        return index < 0 ? CellValue.Error(ErrorKind.NA) : CellValue.Number(index + 1);
    }

    // Match mode 0 (exact), -1, 1, 2 (wildcards), 3 (regular expression); search mode 1, -1, 2, -2.
    private static bool TryReadModes(FunctionCall call, int first, out int matchMode, out int searchMode, out CellValue error)
    {
        matchMode = 0;
        searchMode = 1;
        var match = call.Integer(first, 0);
        if (match.IsError)
        {
            error = match;
            return false;
        }

        var search = call.Integer(first + 1, 1);
        if (search.IsError)
        {
            error = search;
            return false;
        }

        error = CellValue.Error(ErrorKind.Value);
        if (match.AsNumber() is not (-1 or 0 or 1 or 2 or 3) || search.AsNumber() is not (-2 or -1 or 1 or 2))
            return false;
        matchMode = (int)match.AsNumber();
        searchMode = (int)search.AsNumber();
        return true;
    }
}
