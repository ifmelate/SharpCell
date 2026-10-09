using System;
using System.Collections.Generic;
using System.Globalization;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// DSUM, DCOUNT, DGET and the other database functions: <c>Dxxx(database, field, criteria)</c>.
/// <para>
/// The database is a range whose first row holds field names; every row below is a record. The
/// field is a field name (any case) or a 1-based column number. The criteria range has field
/// names in its first row and conditions below: the conditions of one row must all hold, any
/// row may hold, an empty cell holds always (so an empty row selects every record). Conditions
/// are read like COUNTIF criteria except that plain text means "begins with": <c>East</c> also
/// selects <c>Eastern</c>, <c>=East</c> selects only <c>East</c>.
/// </para>
/// </summary>
internal static class DatabaseFunctions
{
    private const int CancellationCheckInterval = 4096;

    private const string ComputedCriteria =
        "A computed criterion (a formula under a name that is not a field) is taken at its value for the first record, not re-evaluated for each record.";

    private static readonly ArgumentKind[] Kinds = [ArgumentKind.Any, ArgumentKind.Any, ArgumentKind.Any];

    public static void Register(FunctionRegistry registry)
    {
        Add(registry, "DSUM", values => Total(values, (sum, x) => sum + x, 0));
        Add(registry, "DPRODUCT", values => Total(values, (product, x) => product * x, 1));
        Add(registry, "DMAX", values => Total(values, Math.Max, double.NegativeInfinity));
        Add(registry, "DMIN", values => Total(values, Math.Min, double.PositiveInfinity));
        Add(registry, "DAVERAGE", Average);
        Add(registry, "DCOUNT", values => CellValue.Number(Numbers(values).Count), countsRecords: true);
        Add(registry, "DCOUNTA", values => CellValue.Number(values.FindAll(v => v.Kind != CellValueKind.Empty).Count), countsRecords: true);
        Add(registry, "DGET", Get);
        Add(registry, "DVAR", values => Variance(values, sample: true, root: false));
        Add(registry, "DVARP", values => Variance(values, sample: false, root: false));
        Add(registry, "DSTDEV", values => Variance(values, sample: true, root: true));
        Add(registry, "DSTDEVP", values => Variance(values, sample: false, root: true));
    }

    private static void Add(FunctionRegistry registry, string name, Func<List<CellValue>, CellValue> aggregate, bool countsRecords = false) =>
        registry.Add(new FunctionInfo(name, 3, 3, Kinds, call => Run(call, aggregate, countsRecords))
        {
            Status = FunctionStatus.KnownDeviation,
            Deviation = ComputedCriteria,
        });

    /// <summary>
    /// Selects the records and passes the field's values in them to <paramref name="aggregate"/>.
    /// DCOUNT and DCOUNTA may omit the field: then each selected record counts.
    /// </summary>
    private static Operand Run(FunctionCall call, Func<List<CellValue>, CellValue> aggregate, bool countsRecords)
    {
        var context = call.Context;
        if (!ValueGrid.TryCreate(call[0], out var database, out var error)
            || !ValueGrid.TryCreate(call[2], out var criteria, out error))
            return error;
        if (database.Rows < 2)
            return CellValue.Error(ErrorKind.Value);

        int? field = null;
        if (!(countsRecords && call.IsMissing(1)))
        {
            var column = FieldColumn(call, database);
            if (column.IsError)
                return column;
            field = (int)column.AsNumber();
        }

        if (!TryReadConditions(context, database, criteria, out var conditions, out error))
            return error;

        var values = new List<CellValue>();
        var visited = 0;
        foreach (var row in RecordRows(database))
        {
            if (++visited % CancellationCheckInterval == 0)
                context.CancellationToken.ThrowIfCancellationRequested();
            if (Selects(context, database, row, conditions))
                values.Add(field is { } f ? database.Get(row, f, context) : CellValue.Number(1));
        }

        return aggregate(values);
    }

    // The field as a 0-based column: a number counts from 1 (truncated, logical values are
    // numbers), text names a field of the header row.
    private static CellValue FieldColumn(FunctionCall call, ValueGrid database)
    {
        var argument = call[1];
        if (argument.Reference is { } reference && !(reference.IsSingleArea && reference.Areas[0].Area.IsSingleCell))
            return CellValue.Error(ErrorKind.Value);

        var field = Evaluator.ToValue(argument, call.Context);
        switch (field.Kind)
        {
            case CellValueKind.Error:
                return field;
            case CellValueKind.Number:
            case CellValueKind.Boolean:
                var index = Math.Truncate(Coercion.ToNumber(field, call.Context.Culture, call.Context.DateSystem).AsNumber());
                return index >= 1 && index <= database.Columns ? CellValue.Number(index - 1) : CellValue.Error(ErrorKind.Value);
            case CellValueKind.Text:
                var column = FindField(call.Context, database, field);
                return column >= 0 ? CellValue.Number(column) : CellValue.Error(ErrorKind.Value);
            default:
                return CellValue.Error(ErrorKind.Value);
        }
    }

    // The first header cell whose text equals the name, ignoring case: TRUE and 1 are names too.
    private static int FindField(EvaluationContext context, ValueGrid database, CellValue name)
    {
        var text = Coercion.ToText(name, context.Culture);
        if (text.IsError || text.AsText().Length == 0)
            return -1;

        for (var c = 0; c < database.Columns; c++)
        {
            var header = Coercion.ToText(database.Get(0, c, context), context.Culture);
            if (!header.IsError && string.Compare(header.AsText(), text.AsText(), context.Culture, CompareOptions.IgnoreCase) == 0)
                return c;
        }

        return -1;
    }

    /// <summary>One criteria row: the conditions that must all hold, or null if it holds for no record.</summary>
    private sealed record ConditionRow(List<(int Column, Criterion Criterion)>? Conditions);

    private static bool TryReadConditions(EvaluationContext context, ValueGrid database, ValueGrid criteria,
        out List<ConditionRow> rows, out CellValue error)
    {
        rows = [];
        error = default;
        if (criteria.Rows < 2)
        {
            error = CellValue.Error(ErrorKind.Value);
            return false;
        }

        var columns = new int[criteria.Columns];
        for (var c = 0; c < criteria.Columns; c++)
            columns[c] = FindField(context, database, criteria.Get(0, c, context));

        for (var r = 1; r < criteria.Rows; r++)
        {
            var conditions = new List<(int, Criterion)>();
            var possible = true;
            for (var c = 0; c < criteria.Columns; c++)
            {
                var value = criteria.Get(r, c, context);
                if (value.Kind == CellValueKind.Empty || (value.Kind == CellValueKind.Text && value.AsText().Length == 0))
                    continue;

                // Under a name that is not a field the cell is a computed criterion: its value
                // must be TRUE for the record (see ComputedCriteria).
                if (columns[c] < 0)
                    possible &= value.Kind == CellValueKind.Boolean && value.AsBoolean();
                else
                    conditions.Add((columns[c], Parse(context, value)));
            }

            // An empty row selects every record, whatever the other rows say.
            if (possible && conditions.Count == 0)
            {
                rows = [new ConditionRow([])];
                return true;
            }

            rows.Add(new ConditionRow(possible ? conditions : null));
        }

        return true;
    }

    // Plain text that is not a number, logical value or error means "begins with".
    private static Criterion Parse(EvaluationContext context, CellValue value)
    {
        if (value.Kind == CellValueKind.Text)
        {
            var text = value.AsText();
            var plain = text[0] is not ('=' or '<' or '>')
                && Coercion.ToNumber(value, context.Culture, context.DateSystem).IsError
                && !text.Equals("TRUE", StringComparison.OrdinalIgnoreCase)
                && !text.Equals("FALSE", StringComparison.OrdinalIgnoreCase)
                && !ErrorKinds.TryParse(text, out _);
            if (plain)
                value = CellValue.Text(text + "*");
        }

        return Criterion.Parse(value, context.Culture, context.DateSystem);
    }

    // Record rows (offsets below the header) that hold anything. A wholly empty record has
    // nothing to sum, count or average, so whole-column databases are not walked cell by cell.
    private static List<int> RecordRows(ValueGrid database)
    {
        var rows = new List<int>();
        if (!database.IsReference)
        {
            for (var r = 1; r < database.Rows; r++)
                rows.Add(r);
            return rows;
        }

        var occupied = new HashSet<(int, int)>();
        database.AddOccupied(occupied);
        var distinct = new HashSet<int>();
        foreach (var (row, _) in occupied)
        {
            if (row > 0 && distinct.Add(row))
                rows.Add(row);
        }

        rows.Sort();
        return rows;
    }

    private static bool Selects(EvaluationContext context, ValueGrid database, int row, List<ConditionRow> criteria)
    {
        foreach (var criteriaRow in criteria)
        {
            if (criteriaRow.Conditions is not { } conditions)
                continue;

            var all = true;
            foreach (var (column, criterion) in conditions)
            {
                if (!criterion.Matches(database.Get(row, column, context)))
                {
                    all = false;
                    break;
                }
            }

            if (all)
                return true;
        }

        return false;
    }

    private static List<double> Numbers(List<CellValue> values)
    {
        var numbers = new List<double>();
        foreach (var value in values)
        {
            if (value.Kind == CellValueKind.Number)
                numbers.Add(value.AsNumber());
        }

        return numbers;
    }

    // Numbers only; with none the result is 0.
    private static CellValue Total(List<CellValue> values, Func<double, double, double> combine, double seed)
    {
        var numbers = Numbers(values);
        if (numbers.Count == 0)
            return CellValue.Number(0);

        var total = seed;
        foreach (var number in numbers)
            total = combine(total, number);
        return CellValue.Number(total);
    }

    private static CellValue Average(List<CellValue> values)
    {
        var numbers = Numbers(values);
        if (numbers.Count == 0)
            return CellValue.Error(ErrorKind.Div0);

        var sum = 0.0;
        foreach (var number in numbers)
            sum += number;
        return CellValue.Number(sum / numbers.Count);
    }

    // Two passes (mean, then squared deviations) for accuracy.
    private static CellValue Variance(List<CellValue> values, bool sample, bool root)
    {
        var numbers = Numbers(values);
        var n = numbers.Count;
        if (n == 0 || (sample && n == 1))
            return CellValue.Error(ErrorKind.Div0);

        var mean = 0.0;
        foreach (var number in numbers)
            mean += number;
        mean /= n;

        var squares = 0.0;
        foreach (var number in numbers)
            squares += (number - mean) * (number - mean);
        var variance = squares / (sample ? n - 1 : n);
        return CellValue.Number(root ? Math.Sqrt(variance) : variance);
    }

    // Exactly one selected record: its value. None is #VALUE!, several are #NUM!.
    private static CellValue Get(List<CellValue> values) => values.Count switch
    {
        0 => CellValue.Error(ErrorKind.Value),
        1 => values[0],
        _ => CellValue.Error(ErrorKind.Num),
    };
}
