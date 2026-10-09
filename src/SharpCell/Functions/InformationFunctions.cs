using System;
using System.Collections.Generic;
using System.Globalization;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

internal static class InformationFunctions
{
    public static void Register(FunctionRegistry registry)
    {
        registry.Add(new FunctionInfo("ISBLANK", 1, 1, [ArgumentKind.Value], call => CellValue.Boolean(call.Value(0).Kind == CellValueKind.Empty)));
        registry.Add(new FunctionInfo("ISERROR", 1, 1, [ArgumentKind.Value], call => CellValue.Boolean(call.Value(0).IsError)));
        registry.Add(new FunctionInfo("ISERR", 1, 1, [ArgumentKind.Value], call => CellValue.Boolean(call.Value(0) is { IsError: true } v && v.AsError() != ErrorKind.NA)));
        registry.Add(new FunctionInfo("ISNA", 1, 1, [ArgumentKind.Value], call => CellValue.Boolean(call.Value(0) is { IsError: true } v && v.AsError() == ErrorKind.NA)));
        registry.Add(new FunctionInfo("ISLOGICAL", 1, 1, [ArgumentKind.Value], call => CellValue.Boolean(call.Value(0).Kind == CellValueKind.Boolean)));
        registry.Add(new FunctionInfo("ISNUMBER", 1, 1, [ArgumentKind.Value], call => CellValue.Boolean(call.Value(0).Kind == CellValueKind.Number)));
        registry.Add(new FunctionInfo("ISTEXT", 1, 1, [ArgumentKind.Value], call => CellValue.Boolean(call.Value(0).Kind == CellValueKind.Text)));
        registry.Add(new FunctionInfo("ISNONTEXT", 1, 1, [ArgumentKind.Value], call => CellValue.Boolean(call.Value(0).Kind != CellValueKind.Text)));
        registry.Add(new FunctionInfo("ISEVEN", 1, 1, [ArgumentKind.Value], call => Parity(call, odd: false)));
        registry.Add(new FunctionInfo("ISODD", 1, 1, [ArgumentKind.Value], call => Parity(call, odd: true)));
        registry.Add(new FunctionInfo("ISREF", 1, 1, [ArgumentKind.Any], call => CellValue.Boolean(call[0].IsReference)));
        registry.Add(new FunctionInfo("ISFORMULA", 1, 1, [ArgumentKind.Any], IsFormula));
        registry.Add(new FunctionInfo("ISOMITTED", 1, 1, [ArgumentKind.Any], call => CellValue.Boolean(call[0] is { IsReference: false, Value.Kind: CellValueKind.Missing })));
        registry.Add(new FunctionInfo("N", 1, 1, [ArgumentKind.Value], N));
        registry.Add(new FunctionInfo("NA", 0, 0, [], _ => CellValue.Error(ErrorKind.NA)));
        registry.Add(new FunctionInfo("TYPE", 1, 1, [ArgumentKind.ScalarAny], TypeOf));
        registry.Add(new FunctionInfo("ERROR.TYPE", 1, 1, [ArgumentKind.Value], ErrorType));
        registry.Add(new FunctionInfo("SHEET", 0, 1, [ArgumentKind.Any], Sheet));
        registry.Add(new FunctionInfo("SHEETS", 0, 1, [ArgumentKind.Any], Sheets));
        registry.Add(new FunctionInfo("CELL", 1, 2, [ArgumentKind.Value, ArgumentKind.Any], Cell)
        {
            Status = FunctionStatus.KnownDeviation,
            Deviation = "Info types that depend on formatting or the file (color, filename, format, parentheses, prefix, protect, width) give #N/A, "
                + "an omitted reference means the formula's own cell rather than the last changed one, and an address on another sheet has no [workbook] part.",
        });
        registry.Add(new FunctionInfo("INFO", 1, 1, [ArgumentKind.Value], Info)
        {
            Status = FunctionStatus.KnownDeviation,
            Deviation = "Describes a fixed environment (Automatic recalculation, release 16.0, system pcdos); \"directory\" gives #N/A.",
        });
        registry.Add(new FunctionInfo("ROWS", 1, 1, [ArgumentKind.Any], call => Dimension(call, rows: true)));
        registry.Add(new FunctionInfo("COLUMNS", 1, 1, [ArgumentKind.Any], call => Dimension(call, rows: false)));
    }

    // The integer part decides; logical values are #VALUE! even though they are numbers elsewhere.
    private static Operand Parity(FunctionCall call, bool odd)
    {
        var value = call.Value(0);
        if (value.Kind == CellValueKind.Boolean)
            return CellValue.Error(ErrorKind.Value);
        var number = Coercion.ToNumber(value, call.Context.Culture, call.Context.DateSystem);
        if (number.IsError)
            return number;
        var isOdd = Math.Abs(Math.Truncate(number.AsNumber()) % 2) == 1;
        return CellValue.Boolean(isOdd == odd);
    }

    // Reads whether cells hold formulas, not their values, so a cell can ask about itself
    // (ISFORMULA(INDIRECT("C9")) in C9). Cells filled by an array result count as formulas.
    private static Operand IsFormula(FunctionCall call)
    {
        var argument = call[0];
        if (argument.Reference is not { } reference)
            return argument.Value.IsError ? argument.Value : CellValue.Error(ErrorKind.Value);
        if (!reference.IsSingleArea)
            return CellValue.Error(ErrorKind.Value);

        var (sheet, area) = reference.Areas[0];
        if (area.IsSingleCell)
            return HasFormula(call.Context, sheet, area.FirstRow, area.FirstColumn);
        if (area.CellCount > Evaluator.MaxArrayCells)
            return CellValue.Error(ErrorKind.Num);

        var result = new CellValue[area.Rows, area.Columns];
        for (var r = 0; r < area.Rows; r++)
        {
            call.Context.CancellationToken.ThrowIfCancellationRequested();
            for (var c = 0; c < area.Columns; c++)
                result[r, c] = HasFormula(call.Context, sheet, area.FirstRow + r, area.FirstColumn + c);
        }

        return CellValue.Array(result);
    }

    private static CellValue HasFormula(EvaluationContext context, Worksheet sheet, int row, int column)
    {
        var data = sheet.Store.Get(row, column);
        if (data is null)
            return CellValue.False;

        // A spilled cell is only known once its anchor is calculated; reading it waits for that.
        if (data.SpillAnchor is { Data: { IsDirty: true, Formula: not null } })
            context.ReadCell(sheet, row, column);
        return CellValue.Boolean(data.Formula is not null || data.SpillAnchor is not null);
    }

    private static Operand N(FunctionCall call)
    {
        var value = call.Value(0);
        return value.Kind switch
        {
            CellValueKind.Number or CellValueKind.Error => value,
            CellValueKind.Boolean => CellValue.Number(value.AsBoolean() ? 1 : 0),
            _ => CellValue.Number(0),
        };
    }

    // A range is an array (64), not the type of its cells.
    private static Operand TypeOf(FunctionCall call) => CellValue.Number(call.Value(0).Kind switch
    {
        CellValueKind.Text => 2,
        CellValueKind.Boolean => 4,
        CellValueKind.Error => 16,
        CellValueKind.Array => 64,
        CellValueKind.Lambda => 128,
        _ => 1,
    });

    // The codes are the ErrorKind values: #NULL! 1 ... #N/A 7, #SPILL! 9, #CALC! 14.
    private static Operand ErrorType(FunctionCall call)
    {
        var value = call.Value(0);
        return value.IsError ? CellValue.Number((int)value.AsError()) : CellValue.Error(ErrorKind.NA);
    }

    // SHEET([value]): the 1-based position of the formula's sheet, of a reference's sheet, or of
    // the sheet with the given name (#N/A if there is none).
    private static Operand Sheet(FunctionCall call)
    {
        var context = call.Context;
        var sheets = context.Workbook.Sheets;
        if (call.Count == 0 || call.IsMissing(0))
            return context.Sheet is { } own ? CellValue.Number(IndexOf(sheets, own)) : CellValue.Error(ErrorKind.NA);

        var argument = call[0];
        if (argument.Reference is { } reference)
            return CellValue.Number(IndexOf(sheets, reference.Areas[0].Sheet));

        return ArrayMath.Map(argument.Value, value =>
        {
            var name = Coercion.ToText(value, context.Culture);
            if (name.IsError)
                return name;
            return context.Workbook.TryGetSheet(name.AsText(), out var sheet)
                ? CellValue.Number(IndexOf(sheets, sheet!))
                : CellValue.Error(ErrorKind.NA);
        });
    }

    // SHEETS([reference]): the sheets in the workbook, or those a (3D) reference spans.
    private static Operand Sheets(FunctionCall call)
    {
        if (call.Count == 0 || call.IsMissing(0))
            return CellValue.Number(call.Context.Workbook.Sheets.Count);

        var argument = call[0];
        if (argument.Reference is not { } reference)
            return argument.Value.IsError ? argument.Value : CellValue.Error(ErrorKind.Value);

        var distinct = new HashSet<Worksheet>();
        foreach (var (sheet, _) in reference.Areas)
            distinct.Add(sheet);
        return CellValue.Number(distinct.Count);
    }

    private static int IndexOf(IReadOnlyList<Worksheet> sheets, Worksheet sheet)
    {
        for (var i = 0; i < sheets.Count; i++)
        {
            if (sheets[i] == sheet)
                return i + 1;
        }

        return 0;
    }

    // CELL(info_type, [reference]) about the top-left cell of the reference.
    private static Operand Cell(FunctionCall call)
    {
        var context = call.Context;
        var info = call.Text(0);
        if (info.IsError)
            return info;

        Worksheet sheet;
        int row, column;
        if (call.Has(1))
        {
            var argument = call[1];
            if (argument.Reference is not { } reference)
                return argument.Value.IsError ? argument.Value : CellValue.Error(ErrorKind.Value);
            (sheet, var area) = reference.Areas[0];
            (row, column) = (area.FirstRow, area.FirstColumn);
        }
        else
        {
            if (context.Sheet is not { } own)
                return CellValue.Error(ErrorKind.Value);
            (sheet, row, column) = (own, context.Origin.Row, context.Origin.Column);
            context.RecordReference(new Reference(own, Area.Cell(row, column)));
        }

        switch (info.AsText().ToUpperInvariant())
        {
            case "ADDRESS":
                var address = "$" + ColumnNames.ToName(column) + "$" + row.ToString(CultureInfo.InvariantCulture);
                return CellValue.Text(sheet == context.Sheet ? address : sheet.Name + "!" + address);
            case "COL":
                return CellValue.Number(column);
            case "ROW":
                return CellValue.Number(row);
            case "CONTENTS":
                return context.ReadCell(sheet, row, column);
            case "TYPE":
                var value = context.ReadCell(sheet, row, column);
                if (value.Kind == CellValueKind.Empty)
                    return CellValue.Text("b");
                var data = sheet.Store.Get(row, column);
                var isConstant = data is { Formula: null, SpillAnchor: null };
                return CellValue.Text(isConstant && value.Kind == CellValueKind.Text ? "l" : "v");
            case "COLOR" or "FILENAME" or "FORMAT" or "PARENTHESES" or "PREFIX" or "PROTECT" or "WIDTH":
                return CellValue.Error(ErrorKind.NA);
            default:
                return CellValue.Error(ErrorKind.Value);
        }
    }

    private static Operand Info(FunctionCall call)
    {
        var type = call.Text(0);
        if (type.IsError)
            return type;

        return type.AsText().ToUpperInvariant() switch
        {
            "NUMFILE" => CellValue.Number(call.Context.Workbook.Sheets.Count),
            "ORIGIN" => CellValue.Text("$A:$A$1"),
            "OSVERSION" => CellValue.Text("Windows (64-bit) NT 10.00"),
            "RECALC" => CellValue.Text("Automatic"),
            "RELEASE" => CellValue.Text("16.0"),
            "SYSTEM" => CellValue.Text("pcdos"),
            "DIRECTORY" => CellValue.Error(ErrorKind.NA),
            _ => CellValue.Error(ErrorKind.Value),
        };
    }

    private static Operand Dimension(FunctionCall call, bool rows)
    {
        var argument = call[0];
        if (argument.Reference is { } reference)
        {
            if (!reference.IsSingleArea)
                return CellValue.Error(ErrorKind.Ref);
            var area = reference.Areas[0].Area;
            return CellValue.Number(rows ? area.Rows : area.Columns);
        }

        var value = argument.Value;
        if (value.IsError)
            return value;
        if (value.Kind != CellValueKind.Array)
            return CellValue.Number(1);
        return CellValue.Number(value.AsArray().GetLength(rows ? 0 : 1));
    }
}
