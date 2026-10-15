using System;

namespace SharpCell;

/// <summary>How a sheet is turned on the printed page, as in Excel's Page Layout &gt; Orientation.</summary>
public enum PageOrientation
{
    /// <summary>Taller than wide; Excel's default.</summary>
    Portrait,

    /// <summary>Wider than tall.</summary>
    Landscape,
}

/// <summary>
/// The margins of a printed sheet in inches, as in Excel's Page Setup &gt; Margins: the space
/// around the cells, and how far the header and footer are from the top and bottom edges.
/// </summary>
public sealed record PageMargins
{
    private const double Max = 49;   // Excel refuses margins that leave no room on the largest paper

    private readonly double _left, _right, _top, _bottom, _header, _footer;

    /// <summary>Margins of the given sizes in inches.</summary>
    /// <exception cref="ArgumentOutOfRangeException">A margin is negative, not a number, or 49 inches or more.</exception>
    public PageMargins(double left, double right, double top, double bottom, double header, double footer)
    {
        _left = Check(left, nameof(left));
        _right = Check(right, nameof(right));
        _top = Check(top, nameof(top));
        _bottom = Check(bottom, nameof(bottom));
        _header = Check(header, nameof(header));
        _footer = Check(footer, nameof(footer));
    }

    /// <summary>Excel's Normal margins, which a sheet has unless set otherwise: 0.7 inches left and right, 0.75 top and bottom, 0.3 for the header and footer.</summary>
    public static PageMargins Normal { get; } = new(0.7, 0.7, 0.75, 0.75, 0.3, 0.3);

    /// <summary>Excel's Wide margins: an inch on every side, 0.5 inches for the header and footer.</summary>
    public static PageMargins Wide { get; } = new(1, 1, 1, 1, 0.5, 0.5);

    /// <summary>Excel's Narrow margins: 0.25 inches left and right, 0.75 top and bottom, 0.3 for the header and footer.</summary>
    public static PageMargins Narrow { get; } = new(0.25, 0.25, 0.75, 0.75, 0.3, 0.3);

    /// <summary>The left margin in inches.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative, not a number, or 49 inches or more.</exception>
    public double Left
    {
        get => _left;
        init => _left = Check(value, nameof(Left));
    }

    /// <summary>The right margin in inches.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative, not a number, or 49 inches or more.</exception>
    public double Right
    {
        get => _right;
        init => _right = Check(value, nameof(Right));
    }

    /// <summary>The top margin in inches.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative, not a number, or 49 inches or more.</exception>
    public double Top
    {
        get => _top;
        init => _top = Check(value, nameof(Top));
    }

    /// <summary>The bottom margin in inches.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative, not a number, or 49 inches or more.</exception>
    public double Bottom
    {
        get => _bottom;
        init => _bottom = Check(value, nameof(Bottom));
    }

    /// <summary>How far the header is from the top edge, in inches.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative, not a number, or 49 inches or more.</exception>
    public double Header
    {
        get => _header;
        init => _header = Check(value, nameof(Header));
    }

    /// <summary>How far the footer is from the bottom edge, in inches.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative, not a number, or 49 inches or more.</exception>
    public double Footer
    {
        get => _footer;
        init => _footer = Check(value, nameof(Footer));
    }

    private static double Check(double value, string name) =>
        value is >= 0 and < Max ? value : throw new ArgumentOutOfRangeException(name, value, $"A margin is at least 0 and less than {Max} inches.");
}
