using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell;

// Conditional formats: rules that change how cells look by what they hold. SharpCell keeps them
// with the cells, saves them, and works out which cells each one applies to.
public sealed partial class Worksheet
{
    private readonly List<ConditionalRule> _conditionalFormats = [];
    private readonly Dictionary<ConditionalRule, (int Generation, RangeStatistics Statistics)> _statistics = [];

    /// <summary>
    /// The conditional formats in priority order, highest first, each with its ranges. The formulas
    /// of a rule are relative to the first cell of its first range.
    /// </summary>
    public IReadOnlyList<(IReadOnlyList<CellRange> Ranges, ConditionalFormat Format)> ConditionalFormats =>
        [.. _conditionalFormats.Select(rule => ((IReadOnlyList<CellRange>)[.. rule.Areas.Select(area => new CellRange(this, area))], rule.Format))];

    /// <summary>The rules in priority order, with their ranges and parsed formulas.</summary>
    internal IReadOnlyList<ConditionalRule> ConditionalRules => _conditionalFormats;

    /// <summary>
    /// Adds a conditional format to a range, below the rules the sheet has: where several apply to a
    /// cell, those added earlier win. Formulas are relative to the range's top-left cell; for
    /// several ranges, to the first one's.
    /// </summary>
    /// <param name="range">The cells, such as <c>B2:B10</c>, or several ranges separated by spaces, such as <c>B2:B10 D2:D10</c>.</param>
    /// <param name="format">The rule.</param>
    /// <exception cref="ArgumentNullException">The format is null.</exception>
    /// <exception cref="ArgumentException">The range is not an A1 range, or the format is one only read from files
    /// (a colour scale, data bars or an icon set).</exception>
    /// <exception cref="FormulaParseException">A formula cannot be parsed; nothing is added.</exception>
    public void AddConditionalFormat(string range, ConditionalFormat format)
    {
        var areas = Areas(range);
        ArgumentNullException.ThrowIfNull(format);
        if (format.Type is ConditionalFormatType.ColorScale or ConditionalFormatType.DataBar or ConditionalFormatType.IconSet)
            throw new ArgumentException($"{format.Type} rules are only read from files.", nameof(format));
        var origin = new CellAddress(areas[0].FirstRow, areas[0].FirstColumn);
        var (formula1, formula2) = format.HasFormulas
            ? (FormulaParser.Parse(format.Formula1!, origin), format.Formula2 is { } second ? FormulaParser.Parse(second, origin) : null)
            : (Condition(format, origin) is { } condition ? FormulaParser.Parse(condition, origin) : null, null);
        Workbook.ThrowIfInCustomFunction();
        _conditionalFormats.Add(new ConditionalRule(areas, format, formula1, formula2));
    }

    /// <summary>
    /// Takes the cells of a range out of every conditional format, as Excel's Clear Rules from
    /// Selected Cells does; a rule left without cells is removed.
    /// </summary>
    /// <param name="range">The cells, such as <c>A1:Z100</c>, or several ranges separated by spaces.</param>
    /// <exception cref="ArgumentException">The range is not an A1 range.</exception>
    /// <exception cref="NotSupportedException">The range takes the first cell of a rule read from a file whose formula
    /// SharpCell cannot read, so the formula could not be moved to the rule's new first cell.</exception>
    public void ClearConditionalFormats(string range)
    {
        var areas = Areas(range);
        Workbook.ThrowIfInCustomFunction();
        var rules = new List<ConditionalRule>(_conditionalFormats);
        foreach (var area in areas)
            rules = Without(rules, area);
        SetConditionalRules(rules);
    }

    /// <summary>Removes every conditional format of the sheet, as Excel's Clear Rules from Entire Sheet does.</summary>
    public void ClearConditionalFormats()
    {
        Workbook.ThrowIfInCustomFunction();
        _conditionalFormats.Clear();
    }

    /// <summary>Adds a rule as a file has it: a formula that cannot be parsed is kept as text.</summary>
    /// <param name="areas">Its ranges.</param>
    /// <param name="format">The rule; for one with a condition formula of its own (text, blanks, errors, dates), that formula's text in <paramref name="condition"/>.</param>
    /// <param name="condition">The condition formula of a rule that does not show one, starting with <c>=</c>.</param>
    /// <param name="source">What the file had for the rule, kept to write it back.</param>
    internal void AddConditionalRule(IReadOnlyList<Area> areas, ConditionalFormat format, string? condition, object? source)
    {
        var origin = new CellAddress(areas[0].FirstRow, areas[0].FirstColumn);
        var (formula1, formula2) = format.HasFormulas ? (Parse(format.Formula1), Parse(format.Formula2)) : (Parse(condition), null);
        _conditionalFormats.Add(new ConditionalRule([.. areas], format, formula1, formula2, source));

        FormulaNode? Parse(string? text)
        {
            if (text is null)
                return null;
            try
            {
                return FormulaParser.Parse(text, origin);
            }
            catch (FormulaParseException ex)
            {
                return new UnsupportedNode(text[1..], ex.Message);
            }
        }
    }

    internal void SetConditionalRules(List<ConditionalRule> rules)
    {
        _conditionalFormats.Clear();
        _conditionalFormats.AddRange(rules);
    }

    /// <summary>
    /// The formula Excel stores for a rule whose condition it writes itself, for the rule's first
    /// cell; null for rules without one.
    /// </summary>
    internal static string? Condition(ConditionalFormat format, CellAddress origin)
    {
        var cell = origin.ToString();
        var text = format.Text is { } t ? "\"" + t.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"" : null;
        return format.Type switch
        {
            ConditionalFormatType.ContainsText => $"=NOT(ISERROR(SEARCH({text},{cell})))",
            ConditionalFormatType.NotContainsText => $"=ISERROR(SEARCH({text},{cell}))",
            ConditionalFormatType.BeginsWith => $"=LEFT({cell},LEN({text}))={text}",
            ConditionalFormatType.EndsWith => $"=RIGHT({cell},LEN({text}))={text}",
            ConditionalFormatType.Blanks => $"=LEN(TRIM({cell}))=0",
            ConditionalFormatType.NoBlanks => $"=LEN(TRIM({cell}))>0",
            ConditionalFormatType.Errors => $"=ISERROR({cell})",
            ConditionalFormatType.NoErrors => $"=NOT(ISERROR({cell}))",
            ConditionalFormatType.TimePeriod => format.Period switch
            {
                TimePeriod.Today => $"=FLOOR({cell},1)=TODAY()",
                TimePeriod.Yesterday => $"=FLOOR({cell},1)=TODAY()-1",
                TimePeriod.Tomorrow => $"=FLOOR({cell},1)=TODAY()+1",
                TimePeriod.Last7Days => $"=AND(TODAY()-FLOOR({cell},1)<=6,FLOOR({cell},1)<=TODAY())",
                TimePeriod.ThisWeek => $"=AND(TODAY()-ROUNDDOWN({cell},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({cell},0)-TODAY()<=7-WEEKDAY(TODAY()))",
                TimePeriod.LastWeek => $"=AND(TODAY()-ROUNDDOWN({cell},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({cell},0)<(WEEKDAY(TODAY())+7))",
                TimePeriod.NextWeek => $"=AND(ROUNDDOWN({cell},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({cell},0)-TODAY()<(15-WEEKDAY(TODAY())))",
                TimePeriod.ThisMonth => $"=AND(MONTH({cell})=MONTH(TODAY()),YEAR({cell})=YEAR(TODAY()))",
                TimePeriod.LastMonth => $"=AND(MONTH({cell})=MONTH(EDATE(TODAY(),0-1)),YEAR({cell})=YEAR(EDATE(TODAY(),0-1)))",
                _ => $"=AND(MONTH({cell})=MONTH(EDATE(TODAY(),0+1)),YEAR({cell})=YEAR(EDATE(TODAY(),0+1)))",
            },
            _ => null,
        };
    }

    /// <summary>
    /// The rules that apply to a cell now, highest priority first, up to the first one that applies
    /// with Stop If True. Values are those of the last calculation.
    /// </summary>
    internal List<ConditionalFormat> MatchingFormats(int row, int column)
    {
        var matching = new List<ConditionalFormat>();
        foreach (var rule in _conditionalFormats)
        {
            if (!rule.Areas.Exists(area => area.Contains(row, column)) || !Matches(rule, row, column))
                continue;
            matching.Add(rule.Format);
            if (rule.Format.StopIfTrue)
                break;
        }

        return matching;
    }

    private bool Matches(ConditionalRule rule, int row, int column)
    {
        var value = Store.Get(row, column)?.Value ?? CellValue.Empty;
        switch (rule.Format.Type)
        {
            case ConditionalFormatType.CellValue:
                return rule.F1 is not null && Truthy(Evaluate(CellValueCondition(rule), row, column));
            case ConditionalFormatType.Top or ConditionalFormatType.Bottom or ConditionalFormatType.AboveAverage
                or ConditionalFormatType.BelowAverage or ConditionalFormatType.DuplicateValues or ConditionalFormatType.UniqueValues:
                return Statistics(rule).Matches(rule.Format, value);
            case ConditionalFormatType.ColorScale or ConditionalFormatType.DataBar or ConditionalFormatType.IconSet:
                return value.Kind == CellValueKind.Number;
            default:
                return rule.F1 is { } condition and not UnsupportedNode && Truthy(Evaluate(condition, row, column));
        }
    }

    private CellValue Evaluate(FormulaNode node, int row, int column) =>
        Workbook.Calculation.EvaluateDetached(node, this, new CellAddress(row, column), CancellationToken.None);

    // TRUE or a number other than 0 applies a rule; text, errors and blanks do not.
    private static bool Truthy(CellValue value)
    {
        if (value.Kind == CellValueKind.Array)
        {
            var array = value.AsArray();
            value = array.Length > 0 ? array[0, 0] : CellValue.Empty;
        }

        return value.Kind switch
        {
            CellValueKind.Boolean => value.AsBoolean(),
            CellValueKind.Number => value.AsNumber() != 0,
            _ => false,
        };
    }

    // The cell compared with the formulas, as a formula: the cell's value is read as a reference is,
    // so an empty cell compares as 0 or "", as in Excel.
    private static FormulaNode CellValueCondition(ConditionalRule rule)
    {
        var self = new ReferenceNode(null, AreaRef.Cell(new CellRef(AxisRef.Relative(0), AxisRef.Relative(0))));
        var first = rule.F1!;
        var second = rule.F2 ?? first;
        return rule.Format.Operator switch
        {
            ConditionalOperator.Between => Between(self, first, second),
            ConditionalOperator.NotBetween => new FunctionNode("NOT", [Between(self, first, second)]),
            ConditionalOperator.Equal => new BinaryNode(BinaryOperator.Equal, self, first),
            ConditionalOperator.NotEqual => new BinaryNode(BinaryOperator.NotEqual, self, first),
            ConditionalOperator.GreaterThan => new BinaryNode(BinaryOperator.Greater, self, first),
            ConditionalOperator.LessThan => new BinaryNode(BinaryOperator.Less, self, first),
            ConditionalOperator.GreaterThanOrEqual => new BinaryNode(BinaryOperator.GreaterEqual, self, first),
            _ => new BinaryNode(BinaryOperator.LessEqual, self, first),
        };

        // Either way round: the bounds may be given high first.
        static FormulaNode Between(FormulaNode value, FormulaNode a, FormulaNode b) => new FunctionNode("OR",
        [
            new FunctionNode("AND", [new BinaryNode(BinaryOperator.GreaterEqual, value, a), new BinaryNode(BinaryOperator.LessEqual, value, b)]),
            new FunctionNode("AND", [new BinaryNode(BinaryOperator.GreaterEqual, value, b), new BinaryNode(BinaryOperator.LessEqual, value, a)]),
        ]);
    }

    // What ranking and counting rules need of the values of their ranges, worked out once per state of the workbook.
    private RangeStatistics Statistics(ConditionalRule rule)
    {
        var generation = Workbook.Calculation.Generation;
        if (_statistics.TryGetValue(rule, out var cached) && cached.Generation == generation)
            return cached.Statistics;
        var values = new List<CellValue>();
        var seen = new HashSet<(int, int)>();
        foreach (var area in rule.Areas)
        {
            foreach (var cell in Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
            {
                if (seen.Add((cell.Row, cell.Column)))
                    values.Add(cell.Data.Value);
            }
        }

        var statistics = new RangeStatistics(values);
        if (_statistics.Count > 64)
            _statistics.Clear();
        _statistics[rule] = (generation, statistics);
        return statistics;
    }

    // Rows or columns of a sheet were inserted or deleted: as data validation rules do, the rules'
    // cells move on that sheet, and references into the edited sheet follow on every sheet.
    private void FollowEditInFormats(Worksheet edited, SheetEdit edit)
    {
        for (var i = _conditionalFormats.Count - 1; i >= 0; i--)
        {
            var rule = _conditionalFormats[i];
            if (edited != this)
            {
                _conditionalFormats[i] = Shift(rule, rule.Origin, rule.Origin, edited, edit);
                continue;
            }

            var areas = new List<Area>(rule.Areas.Count);
            CellAddress? kept = null;
            foreach (var area in rule.Areas)
            {
                var parts = edit.MapCells(area);
                if (parts.Count == 0)
                    continue;
                kept ??= FirstKept(parts[0].Before, edit);
                areas.AddRange(parts.Select(part => part.After));
            }

            if (areas.Count == 0)
                _conditionalFormats.RemoveAt(i);
            else
                _conditionalFormats[i] = Shift(rule, kept!.Value, new CellAddress(areas[0].FirstRow, areas[0].FirstColumn), edited, edit) with { Areas = areas };
        }
    }

    private ConditionalRule Shift(ConditionalRule rule, CellAddress from, CellAddress to, Worksheet edited, SheetEdit edit)
    {
        var f1 = rule.F1 is null ? null : ReferenceShift.Apply(rule.F1, this, from, to, edited, edit);
        var f2 = rule.F2 is null ? null : ReferenceShift.Apply(rule.F2, this, from, to, edited, edit);
        if (f1 == rule.F1 && f2 == rule.F2 && rule.Origin == to)
            return rule;
        var moved = rule with { F1 = f1, F2 = f2 };
        return moved with { Format = PrintedAt(moved, to) };
    }

    // The rules with an area's cells taken out; a rule losing its first cell has its formulas written for the new one.
    private List<ConditionalRule> Without(List<ConditionalRule> from, Area cleared)
    {
        var rules = new List<ConditionalRule>(from.Count);
        foreach (var rule in from)
        {
            if (!rule.Areas.Exists(area => area.TryIntersect(cleared, out _)))
            {
                rules.Add(rule);
                continue;
            }

            var areas = new List<Area>(rule.Areas.Count + 3);
            foreach (var area in rule.Areas)
                Subtract(area, cleared, areas);
            if (areas.Count == 0)
                continue;
            var origin = new CellAddress(areas[0].FirstRow, areas[0].FirstColumn);
            if (origin != rule.Origin && (rule.F1 is UnsupportedNode || rule.F2 is UnsupportedNode))
                throw new NotSupportedException($"A conditional format on {Name}!{new CellRange(this, rule.Areas[0]).Address} has a formula SharpCell cannot read; it could not follow the rule's new first cell.");
            rules.Add(rule with { Areas = areas, Format = PrintedAt(rule, origin) });
        }

        return rules;
    }

    // The rule's formulas written for another first cell: relative references are offsets, so the
    // nodes stay and only their text changes.
    private static ConditionalFormat PrintedAt(ConditionalRule rule, CellAddress origin) =>
        rule.Format.HasFormulas
            ? rule.Format.WithFormulas(rule.F1 is null ? null : "=" + FormulaPrinter.Print(rule.F1, origin), rule.F2 is null ? null : "=" + FormulaPrinter.Print(rule.F2, origin))
            : rule.Format;

    /// <summary>The rules on the cells of an area, for copying them: their ranges inside the area, counted from its top-left cell.</summary>
    /// <exception cref="NotSupportedException">A rule on the area has a formula SharpCell cannot read.</exception>
    internal List<ConditionalRule> ConditionalRulesIn(Area area)
    {
        var copied = new List<ConditionalRule>();
        foreach (var rule in _conditionalFormats)
        {
            var parts = new List<Area>();
            foreach (var part in rule.Areas)
            {
                if (part.TryIntersect(area, out var inside))
                    parts.Add(new Area(inside.FirstRow - area.FirstRow, inside.FirstColumn - area.FirstColumn, inside.LastRow - area.FirstRow, inside.LastColumn - area.FirstColumn));
            }

            if (parts.Count == 0)
                continue;
            if (rule.F1 is UnsupportedNode || rule.F2 is UnsupportedNode)
                throw new NotSupportedException($"A conditional format on {Name}!{new CellRange(this, rule.Areas[0]).Address} has a formula SharpCell cannot read; it could not be copied.");
            copied.Add(rule with { Areas = parts });
        }

        return copied;
    }

    /// <summary>The sheet's rules after rules are copied onto areas: their cells leave the rules they had and take the copied ones.</summary>
    internal List<ConditionalRule> ConditionalRulesPasted(List<ConditionalRule> copied, IEnumerable<Area> tiles)
    {
        var rules = new List<ConditionalRule>(_conditionalFormats);
        foreach (var tile in tiles)
        {
            rules = Without(rules, tile);
            foreach (var rule in copied)
            {
                var areas = rule.Areas.ConvertAll(a => new Area(a.FirstRow + tile.FirstRow, a.FirstColumn + tile.FirstColumn, a.LastRow + tile.FirstRow, a.LastColumn + tile.FirstColumn));
                var origin = new CellAddress(areas[0].FirstRow, areas[0].FirstColumn);
                var f1 = rule.F1 is null ? null : ReferenceShift.Copy(rule.F1, origin);
                var f2 = rule.F2 is null ? null : ReferenceShift.Copy(rule.F2, origin);
                var pasted = new ConditionalRule(areas, rule.Format, f1, f2);
                rules.Add(pasted with { Format = PrintedAt(pasted, origin) });
            }
        }

        return rules;
    }

    private void CopyConditionalFormatsFrom(Worksheet source)
    {
        foreach (var rule in source._conditionalFormats)
            _conditionalFormats.Add(rule with { Areas = [.. rule.Areas] });
    }
}

/// <summary>
/// A conditional format on a sheet: its ranges, the rule, and its formulas parsed relative to the
/// first cell of the first range (an <see cref="UnsupportedNode"/> for one SharpCell cannot parse).
/// For a cell value rule they are the compared formulas; for text, blank, error and date rules the
/// condition Excel stores.
/// </summary>
/// <param name="Areas">The ranges, the first one's top-left cell first.</param>
/// <param name="Format">The rule.</param>
/// <param name="F1">The first formula.</param>
/// <param name="F2">The second formula of a between rule.</param>
/// <param name="Source">What the file had for the rule, written back for what SharpCell does not model.</param>
internal sealed record ConditionalRule(List<Area> Areas, ConditionalFormat Format, FormulaNode? F1, FormulaNode? F2, object? Source = null)
{
    /// <summary>The cell the formulas are relative to.</summary>
    public CellAddress Origin => new(Areas[0].FirstRow, Areas[0].FirstColumn);
}

/// <summary>What ranking and counting rules know of their ranges' values.</summary>
internal sealed class RangeStatistics
{
    private readonly double[] _numbers;
    private readonly Dictionary<string, int> _counts = new(StringComparer.Ordinal);

    public RangeStatistics(List<CellValue> values)
    {
        _numbers = [.. values.Where(v => v.Kind == CellValueKind.Number).Select(v => v.AsNumber()).Order()];
        foreach (var value in values)
        {
            if (Key(value) is { } key)
                _counts[key] = _counts.TryGetValue(key, out var count) ? count + 1 : 1;
        }

        if (_numbers.Length == 0)
            return;
        Mean = _numbers.Average();
        if (_numbers.Length > 1)
            Deviation = Math.Sqrt(_numbers.Sum(n => (n - Mean) * (n - Mean)) / (_numbers.Length - 1));
    }

    public double Mean { get; }

    // The sample standard deviation, as STDEV gives.
    public double Deviation { get; }

    public bool Matches(ConditionalFormat format, CellValue value)
    {
        switch (format.Type)
        {
            case ConditionalFormatType.DuplicateValues:
                return Key(value) is { } duplicate && _counts[duplicate] > 1;
            case ConditionalFormatType.UniqueValues:
                return Key(value) is { } unique && _counts[unique] == 1;
        }

        if (value.Kind != CellValueKind.Number || _numbers.Length == 0)
            return false;
        var number = value.AsNumber();
        switch (format.Type)
        {
            case ConditionalFormatType.Top or ConditionalFormatType.Bottom:
            {
                var count = format.Percent ? Math.Max(1, (int)Math.Floor(_numbers.Length * format.Rank / 100.0)) : format.Rank;
                count = Math.Min(count, _numbers.Length);
                return format.Type == ConditionalFormatType.Top
                    ? number >= _numbers[^count]
                    : number <= _numbers[count - 1];
            }

            case ConditionalFormatType.AboveAverage:
            {
                var limit = Mean + format.StandardDeviations * Deviation;
                return format.OrEqual ? number >= limit : number > limit;
            }

            default:
            {
                var limit = Mean - format.StandardDeviations * Deviation;
                return format.OrEqual ? number <= limit : number < limit;
            }
        }
    }

    // Values compared as Excel compares them for duplicates: numbers by value, text ignoring case;
    // empty cells take no part.
    private static string? Key(CellValue value) => value.Kind switch
    {
        CellValueKind.Number => "n" + value.AsNumber().ToString("R", CultureInfo.InvariantCulture),
        CellValueKind.Text when value.AsText().Length > 0 => "t" + value.AsText().ToUpperInvariant(),
        CellValueKind.Boolean => "b" + value.AsBoolean(),
        CellValueKind.Error => "e" + value.ToString(),
        _ => null,
    };
}
