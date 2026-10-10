using System;

namespace SharpCell;

/// <summary>The font of a cell's text. Change one with <c>with</c>: <c>font with { Bold = true }</c>.</summary>
public sealed record CellFont
{
    private readonly string _name = "Calibri";
    private readonly double _size = 11;

    /// <summary>Calibri 11 in the automatic colour, the font of a new workbook in Excel 2007 to 2021.</summary>
    public static CellFont Default { get; } = new();

    /// <summary>The font family, such as <c>Calibri</c>.</summary>
    /// <exception cref="ArgumentException">The name is empty.</exception>
    public string Name
    {
        get => _name;
        init
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(value);
            _name = value;
        }
    }

    /// <summary>The size in points, 1 to 409 as in Excel.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The size is outside 1 to 409.</exception>
    public double Size
    {
        get => _size;
        init
        {
            if (!(value is >= 1 and <= 409))
                throw new ArgumentOutOfRangeException(nameof(value), value, "Excel font sizes go from 1 to 409 points.");
            _size = value;
        }
    }

    /// <summary>Bold text.</summary>
    public bool Bold { get; init; }

    /// <summary>Italic text.</summary>
    public bool Italic { get; init; }

    /// <summary>The underline, if any.</summary>
    public CellUnderline Underline { get; init; }

    /// <summary>Text struck through.</summary>
    public bool Strikethrough { get; init; }

    /// <summary>The text colour; null for automatic, which Excel shows black.</summary>
    public CellColor? Color { get; init; }
}

/// <summary>How a font underlines text.</summary>
public enum CellUnderline
{
    /// <summary>No underline.</summary>
    None,

    /// <summary>One line under the characters.</summary>
    Single,

    /// <summary>Two lines under the characters.</summary>
    Double,

    /// <summary>One line across the whole cell, lower than <see cref="Single"/>.</summary>
    SingleAccounting,

    /// <summary>Two lines across the whole cell, lower than <see cref="Double"/>.</summary>
    DoubleAccounting,
}
