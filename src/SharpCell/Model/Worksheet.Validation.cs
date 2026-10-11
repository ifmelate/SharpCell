using System;
using System.Collections.Generic;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell;

// Data validation: rules on what cells accept, kept and saved with the sheet. Calculation never sees them.
public sealed partial class Worksheet
{
    private readonly List<ValidationRule> _validations = [];

    /// <summary>
    /// The data validation rules, one entry per range: a rule over several ranges has an entry for
    /// each, in order. The formulas of a rule are relative to the first cell of its first range.
    /// </summary>
    public IReadOnlyList<(CellRange Range, DataValidation Validation)> Validations
    {
        get
        {
            var list = new List<(CellRange, DataValidation)>();
            foreach (var rule in _validations)
            {
                foreach (var area in rule.Areas)
                    list.Add((new CellRange(this, area), rule.Validation));
            }

            return list;
        }
    }

    /// <summary>The rules in order, with their ranges and parsed formulas.</summary>
    internal IReadOnlyList<ValidationRule> ValidationRules => _validations;

    /// <summary>
    /// Adds a data validation rule to a range. Cells of the range leave the rules they had, as in
    /// Excel, where a cell has one rule. The formulas are relative to the range's top-left cell; for
    /// several ranges, to the first one's.
    /// </summary>
    /// <param name="range">The cells, such as <c>B2:B10</c>, or several ranges separated by spaces as
    /// Excel writes them, such as <c>B2:B10 D2:D10</c>.</param>
    /// <param name="validation">The rule.</param>
    /// <exception cref="ArgumentNullException">The validation is null.</exception>
    /// <exception cref="ArgumentException">The range is not an A1 range; or a rule other than
    /// <see cref="ValidationType.Any"/> has no <see cref="DataValidation.Formula1"/>; or a comparing rule
    /// (whole number, decimal, date, time, text length) with <see cref="ValidationOperator.Between"/> or
    /// <see cref="ValidationOperator.NotBetween"/> has no <see cref="DataValidation.Formula2"/>.</exception>
    /// <exception cref="FormulaParseException">A formula cannot be parsed; nothing is added.</exception>
    /// <exception cref="NotSupportedException">The range takes the first cell of a rule read from a file whose formula
    /// SharpCell cannot read, so the formula could not be moved to the rule's new first cell.</exception>
    public void AddValidation(string range, DataValidation validation)
    {
        var areas = Areas(range);
        var area = areas[0];
        ArgumentNullException.ThrowIfNull(validation);
        if (validation.Type != ValidationType.Any && validation.Formula1 is null)
            throw new ArgumentException($"A {validation.Type} rule needs Formula1.", nameof(validation));
        if (validation.Type is not (ValidationType.Any or ValidationType.List or ValidationType.Custom)
            && validation.Operator is ValidationOperator.Between or ValidationOperator.NotBetween && validation.Formula2 is null)
            throw new ArgumentException($"A {validation.Type} rule with {validation.Operator} needs Formula2.", nameof(validation));
        var origin = new CellAddress(area.FirstRow, area.FirstColumn);
        var formula1 = validation.Formula1 is { } text1 ? FormulaParser.Parse(text1, origin) : null;
        var formula2 = validation.Formula2 is { } text2 ? FormulaParser.Parse(text2, origin) : null;
        Workbook.ThrowIfInCustomFunction();

        var rules = new List<ValidationRule>(_validations);
        foreach (var part in areas)
            rules = Without(rules, part);
        rules.Add(new ValidationRule(areas, validation, formula1, formula2));
        _validations.Clear();
        _validations.AddRange(rules);
    }

    // One range, or several separated by spaces.
    private List<Area> Areas(string range)
    {
        ArgumentNullException.ThrowIfNull(range);
        var parts = range.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0)
            throw new ArgumentException("The range is empty.", nameof(range));
        var areas = new List<Area>(parts.Length);
        foreach (var part in parts)
            areas.Add(Range(part).Area);
        return areas;
    }

    /// <summary>
    /// Takes the cells of a range out of every data validation rule; a rule left without cells is
    /// removed. A rule that loses its first cell keeps its formulas relative to its new first cell,
    /// so every remaining cell is checked as before.
    /// </summary>
    /// <param name="range">The cells, such as <c>A1:Z100</c>, or several ranges separated by spaces.</param>
    /// <exception cref="ArgumentException">The range is not an A1 range.</exception>
    /// <exception cref="NotSupportedException">The range takes the first cell of a rule read from a file whose formula
    /// SharpCell cannot read, so the formula could not be moved to the rule's new first cell.</exception>
    public void ClearValidation(string range)
    {
        var areas = Areas(range);
        Workbook.ThrowIfInCustomFunction();

        var rules = new List<ValidationRule>(_validations);
        foreach (var area in areas)
            rules = Without(rules, area);
        _validations.Clear();
        _validations.AddRange(rules);
    }

    /// <summary>
    /// Adds a rule as a file has it: cells are not taken out of other rules, and a formula that
    /// cannot be parsed is kept as text.
    /// </summary>
    internal void AddValidationRule(IReadOnlyList<Area> areas, DataValidation validation)
    {
        if (areas.Count == 0)
            throw new ArgumentException("A rule needs at least one range.", nameof(areas));
        var origin = new CellAddress(areas[0].FirstRow, areas[0].FirstColumn);
        _validations.Add(new ValidationRule([.. areas], validation, Parse(validation.Formula1), Parse(validation.Formula2)));

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

    internal DataValidation? ValidationAt(int row, int column)
    {
        foreach (var rule in _validations)
        {
            foreach (var area in rule.Areas)
            {
                if (area.Contains(row, column))
                    return rule.Validation;
            }
        }

        return null;
    }

    // The rules with the area's cells taken out, built without changing the sheet, so a rule that
    // cannot be moved stops the change before anything happens.
    private List<ValidationRule> Without(List<ValidationRule> from, Area cleared)
    {
        var rules = new List<ValidationRule>(from.Count + 1);
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
            if (areas.Count > 0)
                rules.Add(Rebase(rule, areas));
        }

        return rules;
    }

    // The parts of an area outside another, in reading order: above, left, right, below.
    private static void Subtract(Area area, Area cleared, List<Area> parts)
    {
        if (!area.TryIntersect(cleared, out var hole))
        {
            parts.Add(area);
            return;
        }

        if (hole.FirstRow > area.FirstRow)
            parts.Add(area with { LastRow = hole.FirstRow - 1 });
        if (hole.FirstColumn > area.FirstColumn)
            parts.Add(new Area(hole.FirstRow, area.FirstColumn, hole.LastRow, hole.FirstColumn - 1));
        if (hole.LastColumn < area.LastColumn)
            parts.Add(new Area(hole.FirstRow, hole.LastColumn + 1, hole.LastRow, area.LastColumn));
        if (hole.LastRow < area.LastRow)
            parts.Add(area with { FirstRow = hole.LastRow + 1 });
    }

    // A rule over other areas. Parsed relative references are offsets from the first cell, so when
    // that cell changes the formulas keep their nodes and are only written again from the new cell:
    // every remaining cell checks the same neighbours as before.
    private ValidationRule Rebase(ValidationRule rule, List<Area> areas)
    {
        var updated = rule with { Areas = areas };
        var origin = updated.Origin;
        if (origin == rule.Origin)
            return updated;
        if (rule.F1 is UnsupportedNode || rule.F2 is UnsupportedNode)
            throw new NotSupportedException($"A validation rule on {Name}!{new CellRange(this, rule.Areas[0]).Address} has a formula SharpCell cannot read; it could not follow the rule's new first cell.");
        return updated with { Validation = rule.Validation.WithFormulas(Print(rule.F1), Print(rule.F2)) };

        string? Print(FormulaNode? node) => node is null ? null : "=" + FormulaPrinter.Print(node, origin);
    }

    // Rows or columns of a sheet were inserted or deleted. On that sheet the rules' cells move, a
    // rule whose cells are all deleted is removed, and the formulas are moved to the rule's new first
    // cell; on every sheet references into the edited sheet follow its cells, as in cell formulas.
    internal void FollowEdit(Worksheet edited, SheetEdit edit)
    {
        for (var i = _validations.Count - 1; i >= 0; i--)
        {
            var rule = _validations[i];
            if (edited != this)
            {
                _validations[i] = Shift(rule, rule.Origin, rule.Origin, edited, edit);
                continue;
            }

            var areas = new List<Area>(rule.Areas.Count);
            CellAddress? kept = null;
            foreach (var area in rule.Areas)
            {
                if (edit.Map(area) is not { } moved)
                    continue;
                kept ??= FirstKept(area, edit);
                areas.Add(moved);
            }

            if (areas.Count == 0)
                _validations.RemoveAt(i);
            else
                _validations[i] = Shift(rule, kept!.Value, new CellAddress(areas[0].FirstRow, areas[0].FirstColumn), edited, edit) with { Areas = areas };
        }
    }

    // The first cell of an area that the edit keeps: the one that becomes the area's first cell.
    private static CellAddress FirstKept(Area area, SheetEdit edit)
    {
        var first = edit.Rows ? area.FirstRow : area.FirstColumn;
        if (edit.Delete && first >= edit.At && first < edit.At + edit.Count)
            first = edit.At + edit.Count;
        return edit.Rows ? new CellAddress(first, area.FirstColumn) : new CellAddress(area.FirstRow, first);
    }

    // The rule's formulas moved by the edit. They are offsets from the rule's first cell, so they are
    // read as the formulas of the cell that becomes the first cell (from), which moves to its new place
    // (to): each remaining cell checks the same cells as before, wherever they went. The text is written
    // again when a formula changed or the rule's first cell is no longer where it was.
    private ValidationRule Shift(ValidationRule rule, CellAddress from, CellAddress to, Worksheet edited, SheetEdit edit)
    {
        var f1 = rule.F1 is null ? null : ReferenceShift.Apply(rule.F1, this, from, to, edited, edit);
        var f2 = rule.F2 is null ? null : ReferenceShift.Apply(rule.F2, this, from, to, edited, edit);
        if (f1 == rule.F1 && f2 == rule.F2 && rule.Origin == to)
            return rule;
        return rule with { Validation = rule.Validation.WithFormulas(Print(f1), Print(f2)), F1 = f1, F2 = f2 };

        string? Print(FormulaNode? node) => node is null ? null : "=" + FormulaPrinter.Print(node, to);
    }

    /// <summary>
    /// The rules on the cells of an area, for copying them: each with the parts of its ranges inside
    /// the area, counted from the area's top-left cell.
    /// </summary>
    /// <exception cref="NotSupportedException">A rule on the area has a formula SharpCell cannot read.</exception>
    internal List<ValidationRule> ValidationsIn(Area area)
    {
        var copied = new List<ValidationRule>();
        foreach (var rule in _validations)
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
                throw new NotSupportedException($"A validation rule on {Name}!{new CellRange(this, rule.Areas[0]).Address} has a formula SharpCell cannot read; it could not be copied.");
            copied.Add(rule with { Areas = parts });
        }

        return copied;
    }

    /// <summary>
    /// The sheet's rules after rules are copied onto areas: their cells leave the rules they had and
    /// take the copied ones, whose formulas keep their offsets from each cell, as cell formulas do.
    /// </summary>
    /// <param name="copied">The rules from <see cref="ValidationsIn"/>.</param>
    /// <param name="tiles">The areas they are copied onto.</param>
    /// <exception cref="NotSupportedException">A rule losing its first cell has a formula SharpCell cannot read.</exception>
    internal List<ValidationRule> ValidationsPasted(List<ValidationRule> copied, IEnumerable<Area> tiles)
    {
        var rules = new List<ValidationRule>(_validations);
        foreach (var tile in tiles)
        {
            rules = Without(rules, tile);
            foreach (var rule in copied)
            {
                var areas = rule.Areas.ConvertAll(a => new Area(a.FirstRow + tile.FirstRow, a.FirstColumn + tile.FirstColumn, a.LastRow + tile.FirstRow, a.LastColumn + tile.FirstColumn));
                var origin = new CellAddress(areas[0].FirstRow, areas[0].FirstColumn);
                var f1 = rule.F1 is null ? null : ReferenceShift.Copy(rule.F1, origin);
                var f2 = rule.F2 is null ? null : ReferenceShift.Copy(rule.F2, origin);
                rules.Add(new ValidationRule(areas, rule.Validation.WithFormulas(Print(f1), Print(f2)), f1, f2));

                string? Print(FormulaNode? node) => node is null ? null : "=" + FormulaPrinter.Print(node, origin);
            }
        }

        return rules;
    }

    /// <summary>Puts rules from <see cref="ValidationsPasted"/> in place of the sheet's.</summary>
    internal void SetValidationRules(List<ValidationRule> rules)
    {
        _validations.Clear();
        _validations.AddRange(rules);
    }

    private void CopyValidationsFrom(Worksheet source)
    {
        foreach (var rule in source._validations)
            _validations.Add(rule with { Areas = [.. rule.Areas] });
    }
}

/// <summary>
/// A data validation rule on a sheet: its ranges, its settings, and its formulas parsed relative to
/// the first cell of the first range (an <see cref="UnsupportedNode"/> for a formula read from a file
/// that SharpCell cannot parse).
/// </summary>
internal sealed record ValidationRule(List<Area> Areas, DataValidation Validation, FormulaNode? F1, FormulaNode? F2)
{
    /// <summary>The cell the formulas are relative to.</summary>
    public CellAddress Origin => new(Areas[0].FirstRow, Areas[0].FirstColumn);
}
