using System;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell;

/// <summary>
/// A handle to one cell position. Creating a handle does not create the cell; storage is
/// allocated only when a value or formula is set.
/// </summary>
public sealed class Cell
{
    internal Cell(Worksheet worksheet, int row, int column)
    {
        Worksheet = worksheet;
        Row = row;
        Column = column;
    }

    /// <summary>The sheet the cell belongs to.</summary>
    public Worksheet Worksheet { get; }

    /// <summary>The row number, from 1.</summary>
    public int Row { get; }

    /// <summary>The column number, from 1 (A is 1).</summary>
    public int Column { get; }

    /// <summary>The A1 address without the sheet name, such as <c>B2</c>.</summary>
    public string Address => new CellAddress(Row, Column).ToString();

    /// <summary>
    /// The cell's <see cref="CellValue"/>: a constant, or the formula's result as of the last
    /// calculation (values do not change until <see cref="Workbook.Recalculate"/>). Setting a value removes
    /// the formula. Cells hold scalars only: Missing, arrays and lambdas cannot be stored.
    /// </summary>
    /// <exception cref="InvalidOperationException">The cell is part of an array formula other than its top-left cell.</exception>
    public CellValue Value
    {
        get => Worksheet.Store.Get(Row, Column)?.Value ?? CellValue.Empty;
        set
        {
            if (value.Kind is CellValueKind.Missing or CellValueKind.Array or CellValueKind.Lambda)
                throw new ArgumentException($"A cell cannot hold a {value.Kind} value.", nameof(value));

            Worksheet.Workbook.ThrowIfInCustomFunction();
            EnsureNotArrayMember();
            var key = Key;
            var calculation = Worksheet.Workbook.Calculation;
            calculation.BeforeChange(key, Worksheet.Store.Get(Row, Column));
            CellData? data = null;
            if (value.Kind == CellValueKind.Empty)
            {
                Worksheet.Store.Remove(Row, Column);
            }
            else
            {
                data = Worksheet.Store.GetOrCreate(Row, Column);
                data.Value = value;
                data.FormulaText = null;
                data.Formula = null;
                data.IsDirty = false;
                data.IsLegacy = false;
                data.FixedArray = null;

                // Typing into a spilled cell makes it the user's: it now blocks the spill.
                data.SpillAnchor = null;
            }

            calculation.AfterChange(key, data);
        }
    }

    /// <summary>
    /// Formula text starting with '=' (added if missing), or null for no formula. The text is
    /// parsed on assignment; invalid text throws <see cref="FormulaParseException"/> and leaves the cell unchanged.
    /// </summary>
    /// <exception cref="InvalidOperationException">The cell is part of an array formula other than its top-left cell.</exception>
    public string? Formula
    {
        get => Worksheet.Store.Get(Row, Column)?.FormulaText;
        set => SetFormula(value, legacy: false);
    }

    /// <param name="legacy">
    /// The formula predates dynamic arrays (as marked in xlsx files): an array or range result is
    /// reduced to one value by implicit intersection instead of spilling.
    /// </param>
    /// <param name="value">Formula text, or null to remove the formula.</param>
    internal void SetFormula(string? value, bool legacy)
    {
        Worksheet.Workbook.ThrowIfInCustomFunction();
        EnsureNotArrayMember();
        var key = Key;
        var calculation = Worksheet.Workbook.Calculation;
        if (string.IsNullOrEmpty(value))
        {
            var existing = Worksheet.Store.Get(Row, Column);
            if (existing?.FormulaText is not null)
            {
                calculation.BeforeChange(key, existing);
                Worksheet.Store.Remove(Row, Column);
                calculation.AfterChange(key, null);
            }

            return;
        }

        var text = value.StartsWith('=') ? value : "=" + value;
        var node = FormulaParser.Parse(text, new CellAddress(Row, Column));
        calculation.BeforeChange(key, Worksheet.Store.Get(Row, Column));
        var data = Worksheet.Store.GetOrCreate(Row, Column);
        data.FormulaText = text;
        data.Formula = node;
        data.IsLegacy = legacy;
        data.FixedArray = null;
        data.Value = CellValue.Empty;
        data.SpillAnchor = null;
        calculation.AfterChange(key, data);
    }

    private CellKey Key => new(Worksheet, Row, Column);

    // As in Excel, an array formula is changed as a whole, from its top-left cell.
    private void EnsureNotArrayMember()
    {
        if (Worksheet.Workbook.Calculation.FixedArrayAt(Worksheet, Row, Column) is { } anchor)
            throw new InvalidOperationException($"{this} is part of the array formula at {anchor}; change the whole array there.");
    }

    /// <summary>The address with the sheet name, such as <c>Sheet1!B2</c>.</summary>
    public override string ToString() => $"{Worksheet.Name}!{Address}";
}
