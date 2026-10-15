namespace SharpCell;

/// <summary>
/// What a conditional format changes about a cell's look, as Excel's Format Cells for a rule sets
/// it: the font's style and colour, the fill, the borders and the number format. A property left
/// null keeps what the cell has. A record: <c>with</c> makes a changed copy.
/// </summary>
public sealed record ConditionalStyle
{
    /// <summary>A look that changes nothing; set properties in an initializer.</summary>
    public ConditionalStyle()
    {
    }

    /// <summary>Bold, or not bold; null keeps the cell's.</summary>
    public bool? Bold { get; init; }

    /// <summary>Italic, or not italic; null keeps the cell's.</summary>
    public bool? Italic { get; init; }

    /// <summary>The underline; null keeps the cell's.</summary>
    public CellUnderline? Underline { get; init; }

    /// <summary>Struck through, or not; null keeps the cell's.</summary>
    public bool? Strikethrough { get; init; }

    /// <summary>The text colour; null keeps the cell's.</summary>
    public CellColor? FontColor { get; init; }

    /// <summary>The fill colour; null keeps the cell's.</summary>
    public CellColor? Fill { get; init; }

    /// <summary>The left border; null keeps the cell's.</summary>
    public CellBorder? LeftBorder { get; init; }

    /// <summary>The right border; null keeps the cell's.</summary>
    public CellBorder? RightBorder { get; init; }

    /// <summary>The top border; null keeps the cell's.</summary>
    public CellBorder? TopBorder { get; init; }

    /// <summary>The bottom border; null keeps the cell's.</summary>
    public CellBorder? BottomBorder { get; init; }

    /// <summary>The number format code, such as <c>0.00%</c>; null keeps the cell's.</summary>
    public string? NumberFormat { get; init; }

    /// <summary>A light red fill with dark red text, Excel's first preset.</summary>
    public static ConditionalStyle LightRedFill { get; } = new() { Fill = CellColor.FromRgb(0xFFC7CE), FontColor = CellColor.FromRgb(0x9C0006) };

    /// <summary>A yellow fill with dark yellow text.</summary>
    public static ConditionalStyle YellowFill { get; } = new() { Fill = CellColor.FromRgb(0xFFEB9C), FontColor = CellColor.FromRgb(0x9C5700) };

    /// <summary>A green fill with dark green text.</summary>
    public static ConditionalStyle GreenFill { get; } = new() { Fill = CellColor.FromRgb(0xC6EFCE), FontColor = CellColor.FromRgb(0x006100) };

    /// <summary>The style with what this one sets put over it.</summary>
    /// <param name="style">The cell's own style.</param>
    public CellStyle ApplyTo(CellStyle style)
    {
        System.ArgumentNullException.ThrowIfNull(style);
        var font = style.Font with
        {
            Bold = Bold ?? style.Font.Bold,
            Italic = Italic ?? style.Font.Italic,
            Underline = Underline ?? style.Font.Underline,
            Strikethrough = Strikethrough ?? style.Font.Strikethrough,
            Color = FontColor ?? style.Font.Color,
        };
        return style with
        {
            Font = font,
            Fill = Fill ?? style.Fill,
            LeftBorder = LeftBorder ?? style.LeftBorder,
            RightBorder = RightBorder ?? style.RightBorder,
            TopBorder = TopBorder ?? style.TopBorder,
            BottomBorder = BottomBorder ?? style.BottomBorder,
        };
    }

    // Where both set something, this one's wins.
    internal ConditionalStyle Over(ConditionalStyle lower) => new()
    {
        Bold = Bold ?? lower.Bold,
        Italic = Italic ?? lower.Italic,
        Underline = Underline ?? lower.Underline,
        Strikethrough = Strikethrough ?? lower.Strikethrough,
        FontColor = FontColor ?? lower.FontColor,
        Fill = Fill ?? lower.Fill,
        LeftBorder = LeftBorder ?? lower.LeftBorder,
        RightBorder = RightBorder ?? lower.RightBorder,
        TopBorder = TopBorder ?? lower.TopBorder,
        BottomBorder = BottomBorder ?? lower.BottomBorder,
        NumberFormat = NumberFormat ?? lower.NumberFormat,
    };
}
