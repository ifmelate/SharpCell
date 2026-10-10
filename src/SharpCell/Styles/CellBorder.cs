namespace SharpCell;

/// <summary>One edge of a cell's border.</summary>
/// <param name="Style">The line.</param>
/// <param name="Color">The colour; null for automatic, which Excel shows black.</param>
public sealed record CellBorder(CellBorderStyle Style, CellColor? Color = null);

/// <summary>The lines Excel draws borders with.</summary>
public enum CellBorderStyle
{
    /// <summary>A thin solid line.</summary>
    Thin,

    /// <summary>A solid line of medium weight.</summary>
    Medium,

    /// <summary>A thick solid line.</summary>
    Thick,

    /// <summary>Two thin lines.</summary>
    Double,

    /// <summary>A hairline, thinner than <see cref="Thin"/>.</summary>
    Hair,

    /// <summary>A thin line of dots.</summary>
    Dotted,

    /// <summary>A thin dashed line.</summary>
    Dashed,

    /// <summary>A thin line of dashes and dots.</summary>
    DashDot,

    /// <summary>A thin line of a dash and two dots.</summary>
    DashDotDot,

    /// <summary>A dashed line of medium weight.</summary>
    MediumDashed,

    /// <summary>A line of dashes and dots of medium weight.</summary>
    MediumDashDot,

    /// <summary>A line of a dash and two dots of medium weight.</summary>
    MediumDashDotDot,

    /// <summary>A slanted line of dashes and dots of medium weight.</summary>
    SlantDashDot,
}
