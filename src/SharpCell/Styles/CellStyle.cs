using System;

namespace SharpCell;

/// <summary>
/// How a cell looks in Excel: font, fill, borders, alignment and protection (locked, formula hidden). The number format is not part of
/// it; see <see cref="Cell.NumberFormat"/>. Styles play no part in calculation. A record: styles
/// compare by value, and <c>with</c> makes a changed copy, such as
/// <c>style with { Fill = CellColor.FromRgb(0xFFFF00) }</c>.
/// </summary>
public sealed record CellStyle
{
    private readonly CellFont _font = CellFont.Default;
    private readonly int _indent;
    private readonly int _textRotation;
    private readonly string? _name;

    /// <summary>A style like <see cref="Default"/>; set properties in an initializer.</summary>
    public CellStyle()
    {
    }

    /// <summary>The style of a new workbook's cells: <see cref="CellFont.Default"/>, no fill, no borders, General alignment at the bottom.</summary>
    public static CellStyle Default { get; } = new();

    /// <summary>
    /// The named cell style this style is based on, as Excel's Home &gt; Cell Styles applies them,
    /// such as <c>Good</c> or <c>Heading 1</c>; null for Normal. The look is this record's own: take
    /// one from <see cref="Workbook.NamedStyles"/> to give a cell a named style's look and name.
    /// </summary>
    /// <exception cref="ArgumentException">The name is empty or only white space.</exception>
    public string? Name
    {
        get => _name;
        init
        {
            if (value is not null && string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A style's name cannot be empty.", nameof(value));
            _name = value is null || value.Equals("Normal", StringComparison.OrdinalIgnoreCase) ? null : value;
        }
    }

    /// <summary>The font.</summary>
    public CellFont Font
    {
        get => _font;
        init => _font = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>
    /// The colour the cell is filled with; null for none. With a <see cref="FillPattern"/> other than
    /// solid it is the background the pattern is drawn on, as Excel's Format Cells &gt; Fill &gt;
    /// Background Color is.
    /// </summary>
    public CellColor? Fill { get; init; }

    /// <summary>
    /// The pattern of the fill, as Excel's Format Cells &gt; Fill &gt; Pattern Style sets it;
    /// <see cref="CellFillPattern.Solid"/> unless set otherwise, which fills the cell with <see cref="Fill"/>.
    /// </summary>
    public CellFillPattern FillPattern { get; init; }

    /// <summary>The colour a pattern other than solid is drawn in; null for automatic, which Excel shows black.</summary>
    public CellColor? PatternColor { get; init; }

    /// <summary>A gradient the cell is filled with instead of <see cref="Fill"/>; null for none.</summary>
    public CellGradient? Gradient { get; init; }

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

    /// <summary>
    /// Whether the cell is locked: it cannot be changed while its sheet is protected
    /// (see <see cref="Worksheet.Protect"/>). Excel's default is true; set false for input cells.
    /// </summary>
    public bool Locked { get; init; } = true;

    /// <summary>Whether the cell's formula is hidden from the formula bar while its sheet is protected.</summary>
    public bool FormulaHidden { get; init; }

    /// <summary>
    /// The angle of the text in degrees, as Format Cells &gt; Alignment &gt; Orientation sets it:
    /// positive turns it counter-clockwise (up), negative clockwise (down); 0 unless set otherwise.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The angle is outside -90 to 90.</exception>
    public int TextRotation
    {
        get => _textRotation;
        init
        {
            if (value is < -90 or > 90)
                throw new ArgumentOutOfRangeException(nameof(value), value, "Excel turns text from -90 to 90 degrees.");
            _textRotation = value;
        }
    }

    /// <summary>Text written top to bottom, one character under another, as Excel's vertical Text orientation; <see cref="TextRotation"/> then does not apply.</summary>
    public bool VerticalText { get; init; }

    /// <summary>Text made smaller until it fits the cell's width, as Format Cells &gt; Alignment &gt; Shrink to fit.</summary>
    public bool ShrinkToFit { get; init; }

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
