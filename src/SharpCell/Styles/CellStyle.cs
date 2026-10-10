using System;

namespace SharpCell;

/// <summary>
/// How a cell looks in Excel: font, fill, borders and alignment. The number format is not part of
/// it; see <see cref="Cell.NumberFormat"/>. Styles play no part in calculation. A record: styles
/// compare by value, and <c>with</c> makes a changed copy, such as
/// <c>style with { Fill = CellColor.FromRgb(0xFFFF00) }</c>.
/// </summary>
public sealed record CellStyle
{
    private readonly CellFont _font = CellFont.Default;
    private readonly int _indent;

    /// <summary>A style like <see cref="Default"/>; set properties in an initializer.</summary>
    public CellStyle()
    {
    }

    /// <summary>The style of a new workbook's cells: <see cref="CellFont.Default"/>, no fill, no borders, General alignment at the bottom.</summary>
    public static CellStyle Default { get; } = new();

    /// <summary>The font.</summary>
    public CellFont Font
    {
        get => _font;
        init => _font = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>The colour the cell is filled with; null for none. Only solid fills are read from files.</summary>
    public CellColor? Fill { get; init; }

    /// <summary>The left edge of the border; null for none.</summary>
    public CellBorder? LeftBorder { get; init; }

    /// <summary>The right edge of the border; null for none.</summary>
    public CellBorder? RightBorder { get; init; }

    /// <summary>The top edge of the border; null for none.</summary>
    public CellBorder? TopBorder { get; init; }

    /// <summary>The bottom edge of the border; null for none.</summary>
    public CellBorder? BottomBorder { get; init; }

    /// <summary>
    /// Where the text sits across the cell. <see cref="CellHorizontalAlignment.General"/> is
    /// Excel's default: numbers right, text left, logical values and errors in the middle.
    /// </summary>
    public CellHorizontalAlignment HorizontalAlignment { get; init; }

    /// <summary>Where the text sits from top to bottom of the cell; Excel's default is the bottom.</summary>
    public CellVerticalAlignment VerticalAlignment { get; init; } = CellVerticalAlignment.Bottom;

    /// <summary>Whether text wraps onto more lines inside the cell.</summary>
    public bool WrapText { get; init; }

    /// <summary>The indent, in steps of the width of three characters of the default font; 0 to 250.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The indent is outside 0 to 250.</exception>
    public int Indent
    {
        get => _indent;
        init
        {
            if (value is < 0 or > 250)
                throw new ArgumentOutOfRangeException(nameof(value), value, "Excel indents go from 0 to 250.");
            _indent = value;
        }
    }
}

/// <summary>Where a cell's text sits across the cell.</summary>
public enum CellHorizontalAlignment
{
    /// <summary>Excel's default: numbers right, text left, logical values and errors in the middle.</summary>
    General,

    /// <summary>Against the left edge, after the indent.</summary>
    Left,

    /// <summary>In the middle.</summary>
    Center,

    /// <summary>Against the right edge, before the indent.</summary>
    Right,

    /// <summary>The text repeated to fill the cell.</summary>
    Fill,

    /// <summary>Wrapped, each line but the last spread to both edges.</summary>
    Justify,

    /// <summary>In the middle of this cell and the empty cells to its right.</summary>
    CenterContinuous,

    /// <summary>Each line spread evenly across the cell.</summary>
    Distributed,
}

/// <summary>Where a cell's text sits from top to bottom of the cell.</summary>
public enum CellVerticalAlignment
{
    /// <summary>Against the top edge.</summary>
    Top,

    /// <summary>In the middle.</summary>
    Center,

    /// <summary>Against the bottom edge, Excel's default.</summary>
    Bottom,

    /// <summary>Lines spread to the top and bottom edges.</summary>
    Justify,

    /// <summary>Lines spread evenly from top to bottom.</summary>
    Distributed,
}
