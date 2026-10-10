using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
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

    /// <summary>
    /// The cells and ranges the cell's formula read in its last calculation: references in the
    /// formula, in the defined names it uses, and those built while calculating (<c>INDIRECT</c>,
    /// <c>OFFSET</c>). A branch of <c>IF</c> that was not taken is not included. A cell filled by
    /// another cell's spill has that cell as its precedent; a constant or empty cell has none. The
    /// ranges come in the order of their sheets in the workbook, then by their top-left cell.
    /// </summary>
    /// <exception cref="InvalidOperationException">The formula is out of date: call <see cref="Workbook.Recalculate"/> first.</exception>
    public IReadOnlyList<CellRange> Precedents
    {
        get
        {
            var data = Worksheet.Store.Get(Row, Column);
            if (data?.Formula is null)
            {
                return data?.SpillAnchor is { } anchor
                    ? [new CellRange(anchor.Sheet, Area.Cell(anchor.Row, anchor.Column))]
                    : [];
            }

            if (data.IsDirty)
                throw new InvalidOperationException($"{this} has not been calculated since it or its inputs changed; call Recalculate first.");

            var order = Worksheet.Workbook.Calculation.SheetOrder();
            var areas = new List<SheetArea>(data.Registered?.Areas ?? []);
            areas.Sort((a, b) =>
            {
                var bySheet = order[a.Sheet].CompareTo(order[b.Sheet]);
                if (bySheet != 0)
                    return bySheet;
                var x = a.Area;
                var y = b.Area;
                return x.FirstRow != y.FirstRow ? x.FirstRow.CompareTo(y.FirstRow)
                    : x.FirstColumn != y.FirstColumn ? x.FirstColumn.CompareTo(y.FirstColumn)
                    : x.LastRow != y.LastRow ? x.LastRow.CompareTo(y.LastRow)
                    : x.LastColumn.CompareTo(y.LastColumn);
            });
            return areas.ConvertAll(a => new CellRange(a.Sheet, a.Area));
        }
    }

    /// <summary>
    /// The formula cells that read this cell in their last calculation, alone or in a range; for a
    /// cell whose formula spills, also the cells it spilled into. Direct readers only: follow
    /// <see cref="Dependents"/> again for the cells that read those. In the order of their sheets
    /// in the workbook, then by row and column.
    /// </summary>
    /// <exception cref="InvalidOperationException">Some formula of the workbook is out of date: call
    /// <see cref="Workbook.Recalculate"/> first.</exception>
    public IReadOnlyList<Cell> Dependents
    {
        get
        {
            var calculation = Worksheet.Workbook.Calculation;
            if (calculation.HasDirty)
                throw new InvalidOperationException("Formulas of the workbook have not been calculated since they or their inputs changed; call Recalculate first.");

            var key = Key;
            var readers = new HashSet<CellKey>();
            calculation.Graph.ForEachReader(key, reader => readers.Add(reader));
            if (Worksheet.Store.Get(Row, Column)?.SpillArea is { } spill)
            {
                foreach (var cell in Worksheet.Store.Enumerate(spill.FirstRow, spill.FirstColumn, spill.LastRow, spill.LastColumn))
                {
                    if (cell.Data.SpillAnchor == key)
                        readers.Add(new CellKey(Worksheet, cell.Row, cell.Column));
                }
            }

            return calculation.SortedCells([.. readers]);
        }
    }

    /// <summary>
    /// The number format Excel shows the cell with, as a format code such as <c>0.00</c>,
    /// <c>#,##0</c> or <c>yyyy-mm-dd</c>; <c>General</c> when the cell has none. Formats are read
    /// from .xlsx files and shape <see cref="Text"/>; they play no part in calculation. An empty
    /// cell can have a format, and a format alone does not make a cell hold something.
    /// </summary>
    /// <exception cref="ArgumentException">The code is not a number format.</exception>
    public string NumberFormat
    {
        get => Worksheet.FormatAt(Row, Column)?.FormatCode ?? "General";
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            Worksheet.Workbook.ThrowIfInCustomFunction();
            if (value.Length == 0 || value.Equals("General", StringComparison.OrdinalIgnoreCase))
            {
                Worksheet.SetFormat(Row, Column, null);
                return;
            }

            if (!SharpCell.Functions.NumberFormat.TryParse(value, out var format))
                throw new ArgumentException($"'{value}' is not a number format code.", nameof(value));
            Worksheet.SetFormat(Row, Column, format);
        }
    }

    /// <summary>
    /// How the cell looks: font, fill, borders and alignment. A cell without a style of its own has
    /// <see cref="Workbook.DefaultStyle"/>; setting null gives it back. Styles are read from .xlsx
    /// files; they play no part in calculation, and a style alone does not make a cell hold something.
    /// </summary>
    [AllowNull]
    public CellStyle Style
    {
        get => Worksheet.StyleAt(Row, Column) ?? Worksheet.Workbook.DefaultStyle;
        set
        {
            Worksheet.Workbook.ThrowIfInCustomFunction();
            Worksheet.SetStyle(Row, Column, value);
        }
    }

    /// <summary>
    /// The cell as Excel shows it: the value formatted by <see cref="NumberFormat"/> with the
    /// workbook's culture and date system. Column width plays no part, so a number that fits no
    /// width is never cut short; a date format on a number that is no date shows <c>#######</c>.
    /// Errors show as <c>#DIV/0!</c> and the like, logical values as <c>TRUE</c> and <c>FALSE</c>,
    /// an empty cell as an empty string.
    /// </summary>
    public string Text
    {
        get
        {
            var value = Value;
            var format = Worksheet.FormatAt(Row, Column);
            var workbook = Worksheet.Workbook;
            return value.Kind switch
            {
                CellValueKind.Number => format is null
                    ? SharpCell.Functions.NumberFormat.FormatGeneral(value.AsNumber(), workbook.Culture)
                    : format.Format(value.AsNumber(), workbook.Culture, workbook.DateSystem) ?? "#######",
                CellValueKind.Text => format is null ? value.AsText() : format.FormatText(value.AsText()),
                CellValueKind.Boolean => value.AsBoolean() ? "TRUE" : "FALSE",
                CellValueKind.Error => value.AsError().ToText(),
                _ => "",
            };
        }
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
