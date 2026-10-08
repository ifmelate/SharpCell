using System;
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

    public Worksheet Worksheet { get; }

    public int Row { get; }

    public int Column { get; }

    public string Address => new CellAddress(Row, Column).ToString();

    /// <summary>
    /// The constant, or the formula's value as of the last calculation. Setting a value removes
    /// the formula. Cells hold scalars only: Missing, arrays and lambdas cannot be stored.
    /// </summary>
    public CellValue Value
    {
        get => Worksheet.Store.Get(Row, Column)?.Value ?? CellValue.Empty;
        set
        {
            if (value.Kind is CellValueKind.Missing or CellValueKind.Array or CellValueKind.Lambda)
                throw new ArgumentException($"A cell cannot hold a {value.Kind} value.", nameof(value));

            if (value.Kind == CellValueKind.Empty)
            {
                Worksheet.Store.Remove(Row, Column);
            }
            else
            {
                var data = Worksheet.Store.GetOrCreate(Row, Column);
                data.Value = value;
                data.FormulaText = null;
                data.Formula = null;
            }

            Worksheet.Workbook.OnCellChanged(Worksheet, Row, Column);
        }
    }

    /// <summary>
    /// Formula text starting with '=' (added if missing), or null for no formula. The text is
    /// parsed on assignment; invalid text throws <see cref="FormulaParseException"/> and leaves the cell unchanged.
    /// </summary>
    public string? Formula
    {
        get => Worksheet.Store.Get(Row, Column)?.FormulaText;
        set
        {
            if (string.IsNullOrEmpty(value))
            {
                if (Worksheet.Store.Get(Row, Column)?.FormulaText is not null)
                {
                    Worksheet.Store.Remove(Row, Column);
                    Worksheet.Workbook.OnCellChanged(Worksheet, Row, Column);
                }

                return;
            }

            var text = value.StartsWith('=') ? value : "=" + value;
            var node = FormulaParser.Parse(text, new CellAddress(Row, Column));
            var data = Worksheet.Store.GetOrCreate(Row, Column);
            data.FormulaText = text;
            data.Formula = node;
            data.Value = CellValue.Empty;
            Worksheet.Workbook.OnCellChanged(Worksheet, Row, Column);
        }
    }

    public override string ToString() => $"{Worksheet.Name}!{Address}";
}
