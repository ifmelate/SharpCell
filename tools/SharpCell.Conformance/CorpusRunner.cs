using System;
using System.Collections.Generic;
using System.Threading;
using SharpCell.Evaluation;
using SharpCell.Parsing;
using SharpCell.Xlsx;

namespace SharpCell.Conformance;

/// <summary>
/// One compared cell: a formula cell or a cell filled by an array formula. <see cref="Widened"/>
/// marks a cell that matches only within a tolerance stated in <c>tests/corpus/overrides.json</c>.
/// </summary>
internal sealed record CellOutcome(
    string Sheet, CellAddress Cell, string? Formula, CellValue Expected, CellValue Actual, bool Passed, IReadOnlyList<string> Functions,
    bool Widened = false)
{
    public string Address => Cell.ToString();
}

internal sealed record FileResult(string File, IReadOnlyList<CellOutcome> Cells, int Skipped, string? Error)
{
    public int Passed
    {
        get
        {
            var passed = 0;
            foreach (var cell in Cells)
            {
                if (cell.Passed)
                    passed++;
            }

            return passed;
        }
    }
}

/// <summary>
/// Loads a workbook Excel calculated, recalculates it and compares every formula result with the
/// value Excel saved. Numbers match within a relative tolerance, errors by kind, text exactly.
/// </summary>
internal static class CorpusRunner
{
    public const double DefaultTolerance = 1e-9;

    private const string MetadataSheet = "METADATA";

    // Random results are compared by kind only; NOW and TODAY too, unless the file says when it was saved.
    private static readonly HashSet<string> RandomFunctions = new(StringComparer.Ordinal) { "RAND", "RANDBETWEEN", "RANDARRAY" };
    private static readonly HashSet<string> ClockFunctions = new(StringComparer.Ordinal) { "NOW", "TODAY" };

    public static FileResult Run(string path, string name, TimeSpan timeout, CorpusOverrides? overrides = null,
        Func<Workbook, Workbook>? afterCalculation = null)
    {
        overrides ??= CorpusOverrides.Empty;
        Workbook workbook;
        try
        {
            workbook = XlsxReader.Load(path);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            return new FileResult(name, [], 0, $"Load failed: {ex.GetType().Name}: {ex.Message}");
        }

        // Excel calculated some files in another locale; a file does not record it.
        if (overrides.CultureOf(name) is { } culture)
            workbook.Culture = culture;

        var tolerance = DefaultTolerance;
        var clockFixed = false;
        if (workbook.TryGetSheet(MetadataSheet, out var metadata))
        {
            if (metadata!["A1"].Value is { Kind: CellValueKind.Number } custom)
                tolerance = custom.AsNumber();
            for (var row = 1; row <= 32; row++)
            {
                if (metadata[row, 1].Value is { Kind: CellValueKind.Text } label && label.AsText() == "NOW"
                    && metadata[row, 2].Value is { Kind: CellValueKind.Number } now)
                {
                    workbook.Clock = new FixedClock(FromSerial(now.AsNumber()));
                    clockFixed = true;
                }
            }
        }

        // Cached values go before calculating, so a formula that is not recalculated cannot pass by
        // being compared with itself.
        var expected = Snapshot(workbook);
        foreach (var (key, _, _, _) in expected)
        {
            if (key.Data is { Formula: not null } data)
                data.Value = CellValue.Empty;
        }

        try
        {
            using var cancellation = new CancellationTokenSource(timeout);
            workbook.Recalculate(cancellation.Token);
        }
        catch (OperationCanceledException)
        {
            return new FileResult(name, [], 0, $"Calculation did not finish within {timeout.TotalSeconds:0} s.");
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            return new FileResult(name, [], 0, $"Calculation failed: {ex.GetType().Name}: {ex.Message}");
        }

        // A round trip compares what a saved and reloaded workbook holds instead of the calculated one.
        var read = workbook;
        if (afterCalculation is not null)
        {
            try
            {
                read = afterCalculation(workbook);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                return new FileResult(name, [], 0, $"Round trip failed: {ex.GetType().Name}: {ex.Message}");
            }
        }

        var cells = new List<CellOutcome>();
        var skipped = 0;
        foreach (var (key, formula, functions, before) in expected)
        {
            // A formula Excel never calculated has nothing to compare with.
            if (before.Kind == CellValueKind.Empty)
            {
                skipped++;
                continue;
            }

            var actual = ReferenceEquals(read, workbook)
                ? key.Data?.Value ?? CellValue.Empty
                : read[key.Sheet.Name][key.Row, key.Column].Value;
            var byKind = functions.Overlaps(RandomFunctions) || (!clockFixed && functions.Overlaps(ClockFunctions));
            var passed = byKind ? before.Kind == actual.Kind : Matches(before, actual, tolerance);
            var widened = !passed && !byKind && overrides.For(name, key.Sheet.Name, key.Address) is { } rule
                && MatchesWithin(before, actual, rule.Relative ?? tolerance, rule.Absolute ?? 0);
            cells.Add(new CellOutcome(key.Sheet.Name, key.Address, formula, before, actual, passed || widened, [.. functions], widened));
        }

        return new FileResult(name, cells, skipped, null);
    }

    // A stated tolerance: numbers within the relative one, or apart by no more than the absolute one.
    private static bool MatchesWithin(CellValue expected, CellValue actual, double relative, double absolute) =>
        Matches(expected, actual, relative)
        || (expected.Kind == CellValueKind.Number && actual.Kind == CellValueKind.Number
            && Math.Abs(expected.AsNumber() - actual.AsNumber()) <= absolute);

    public static bool Matches(CellValue expected, CellValue actual, double tolerance)
    {
        if (expected.Kind != actual.Kind)
            return false;
        if (expected.Kind != CellValueKind.Number)
            return expected.Kind == CellValueKind.Text ? string.Equals(expected.AsText(), actual.AsText(), StringComparison.Ordinal) : expected.Equals(actual);

        var a = expected.AsNumber();
        var b = actual.AsNumber();
        if (a == b)
            return true;
        return Math.Abs(a - b) <= tolerance * Math.Max(Math.Abs(a), Math.Abs(b));
    }

    private static List<(CellKey Key, string? Formula, HashSet<string> Functions, CellValue Value)> Snapshot(Workbook workbook)
    {
        var result = new List<(CellKey, string?, HashSet<string>, CellValue)>();
        foreach (var sheet in workbook.Sheets)
        {
            foreach (var cell in sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
            {
                var key = new CellKey(sheet, cell.Row, cell.Column);
                if (cell.Data.Formula is { } formula)
                {
                    result.Add((key, cell.Data.FormulaText, FunctionsOf(formula, workbook, sheet), cell.Data.Value));
                }
                else if (cell.Data.SpillAnchor is { } anchor && anchor.Data?.Formula is { } anchorFormula)
                {
                    // A spilled cell counts as a case of the functions its anchor uses.
                    result.Add((key, null, FunctionsOf(anchorFormula, workbook, sheet), cell.Data.Value));
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Names of the functions a formula calls, upper case, without file prefixes. Calls of names the
    /// formula binds itself (LET variables, LAMBDA parameters) are not functions and are left out.
    /// </summary>
    public static HashSet<string> FunctionsOf(FormulaNode root)
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        var bound = new HashSet<string>(StringComparer.Ordinal);
        var stack = new Stack<FormulaNode>();
        stack.Push(root);
        while (stack.Count > 0)
        {
            switch (stack.Pop())
            {
                case FunctionNode f:
                    var name = f.Name.ToUpperInvariant();
                    names.Add(name);
                    for (var i = 0; i < f.Arguments.Count - 1; i++)
                    {
                        if ((name == "LAMBDA" || (name == "LET" && i % 2 == 0)) && f.Arguments[i] is NameNode { Sheet: null } parameter)
                            bound.Add(parameter.Name.ToUpperInvariant());
                    }

                    foreach (var argument in f.Arguments)
                        stack.Push(argument);
                    break;
                case CallNode c:
                    stack.Push(c.Callee);
                    foreach (var argument in c.Arguments)
                        stack.Push(argument);
                    break;
                case BinaryNode b:
                    stack.Push(b.Left);
                    stack.Push(b.Right);
                    break;
                case UnaryNode u:
                    stack.Push(u.Operand);
                    break;
                case ParenthesesNode p:
                    stack.Push(p.Inner);
                    break;
                case SpillNode s:
                    stack.Push(s.Operand);
                    break;
                case ImplicitIntersectionNode i:
                    stack.Push(i.Operand);
                    break;
            }
        }

        names.ExceptWith(bound);
        return names;
    }

    // Defined names holding a LAMBDA are called like functions but are the workbook's own.
    private static HashSet<string> FunctionsOf(FormulaNode formula, Workbook workbook, Worksheet sheet)
    {
        var names = FunctionsOf(formula);
        names.RemoveWhere(name => !Functions.FunctionRegistry.Default.TryGet(name, out _)
            && (workbook.Names.TryGet(name, sheet, out _) || workbook.Names.TryGet(name, null, out _)));
        return names;
    }

    // Serial numbers from 61 on (1900-03-01) are a plain day count from 1899-12-30.
    private static DateTime FromSerial(double serial) => new DateTime(1899, 12, 30).AddDays(serial);

    /// <summary>The moment the file was saved, as local time in a UTC "zone", so NOW returns exactly that serial.</summary>
    private sealed class FixedClock(DateTime local) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => new(DateTime.SpecifyKind(local, DateTimeKind.Utc));

        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }
}
