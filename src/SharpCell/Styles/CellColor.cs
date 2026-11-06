using System;
using System.Globalization;

namespace SharpCell;

/// <summary>
/// A colour as Excel shows it, in sRGB. A colour of the workbook's theme also knows its place in
/// the theme (<see cref="Theme"/>) and how much lighter or darker it is (<see cref="Tint"/>), and is
/// saved that way, so it follows the theme as in Excel; make one with <see cref="Workbook.ThemeColor"/>.
/// Indexed colours of a file are turned into plain RGB. Excel ignores the alpha channel, so there is
/// none. Colours compare by value, the theme and tint included: a theme colour is not equal to the
/// plain colour of the same RGB.
/// </summary>
/// <param name="R">Red, 0 to 255.</param>
/// <param name="G">Green, 0 to 255.</param>
/// <param name="B">Blue, 0 to 255.</param>
public readonly record struct CellColor(byte R, byte G, byte B)
{
    private readonly double _tint;

    /// <summary>The theme colour this is, or null for a plain colour.</summary>
    public ThemeColor? Theme { get; init; }

    /// <summary>
    /// How much the theme colour is darkened (from -1, black) or lightened (to 1, white), as
    /// Excel's colour palette shows its shades; 0 for the colour itself and for plain colours.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The tint is outside -1 to 1.</exception>
    public double Tint
    {
        get => _tint;
        init => _tint = value is >= -1 and <= 1 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "A tint goes from -1 to 1.");
    }

    /// <summary>A colour from a <c>0xRRGGBB</c> number.</summary>
    public static CellColor FromRgb(int rgb) => new((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb);

    /// <summary>The colour as <c>#RRGGBB</c>.</summary>
    public override string ToString() => string.Create(CultureInfo.InvariantCulture, $"#{R:X2}{G:X2}{B:X2}");

    /// <summary>A theme colour: the theme's RGB with the tint applied, and the theme and tint kept.</summary>
    internal static CellColor OfTheme(CellColor themeRgb, ThemeColor theme, double tint) =>
        (tint == 0 ? new CellColor(themeRgb.R, themeRgb.G, themeRgb.B) : Tinted(themeRgb, tint)) with { Theme = theme, Tint = tint };

    /// <summary>
    /// A colour with its lightness moved by a tint from -1 (black) to 1 (white), the way OOXML
    /// defines it: the colour in hue, lightness and saturation, the lightness scaled towards 0 or 1.
    /// </summary>
    internal static CellColor Tinted(CellColor color, double tint)
    {
        var (hue, lightness, saturation) = ToHls(color);
        lightness = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
        return FromHls(hue, Math.Clamp(lightness, 0, 1), saturation);
    }

    private static (double Hue, double Lightness, double Saturation) ToHls(CellColor color)
    {
        double r = color.R / 255.0, g = color.G / 255.0, b = color.B / 255.0;
        var max = Math.Max(r, Math.Max(g, b));
        var min = Math.Min(r, Math.Min(g, b));
        var lightness = (max + min) / 2;
        if (max == min)
            return (0, lightness, 0);
        var delta = max - min;
        var saturation = lightness <= 0.5 ? delta / (max + min) : delta / (2 - max - min);
        double hue;
        if (max == r)
            hue = (g - b) / delta;
        else if (max == g)
            hue = 2 + (b - r) / delta;
        else
            hue = 4 + (r - g) / delta;
        hue = (hue / 6) % 1;
        return (hue < 0 ? hue + 1 : hue, lightness, saturation);
    }

    private static CellColor FromHls(double hue, double lightness, double saturation)
    {
        if (saturation == 0)
            return new CellColor(Channel(lightness), Channel(lightness), Channel(lightness));
        var m2 = lightness <= 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation;
        var m1 = 2 * lightness - m2;
        return new CellColor(Channel(Component(m1, m2, hue + 1.0 / 3)), Channel(Component(m1, m2, hue)), Channel(Component(m1, m2, hue - 1.0 / 3)));
    }

    private static double Component(double m1, double m2, double hue)
    {
        hue = hue < 0 ? hue + 1 : hue > 1 ? hue - 1 : hue;
        if (hue < 1.0 / 6)
            return m1 + (m2 - m1) * hue * 6;
        if (hue < 0.5)
            return m2;
        if (hue < 2.0 / 3)
            return m1 + (m2 - m1) * (2.0 / 3 - hue) * 6;
        return m1;
    }

    private static byte Channel(double value) => (byte)Math.Clamp(Math.Floor(value * 255 + 0.5), 0, 255);
}

/// <summary>
/// The twelve colours of a workbook's theme, by the index .xlsx files give them, as Excel's colour
/// palette names them.
/// </summary>
public enum ThemeColor
{
    /// <summary>Background 1, the theme's first light colour (white in the Office theme).</summary>
    Background1 = 0,

    /// <summary>Text 1, the theme's first dark colour (black in the Office theme).</summary>
    Text1 = 1,

    /// <summary>Background 2, the second light colour.</summary>
    Background2 = 2,

    /// <summary>Text 2, the second dark colour.</summary>
    Text2 = 3,

    /// <summary>Accent 1.</summary>
    Accent1 = 4,

    /// <summary>Accent 2.</summary>
    Accent2 = 5,

    /// <summary>Accent 3.</summary>
    Accent3 = 6,

    /// <summary>Accent 4.</summary>
    Accent4 = 7,

    /// <summary>Accent 5.</summary>
    Accent5 = 8,

    /// <summary>Accent 6.</summary>
    Accent6 = 9,

    /// <summary>The colour of hyperlinks.</summary>
    Hyperlink = 10,

    /// <summary>The colour of followed hyperlinks.</summary>
    FollowedHyperlink = 11,
}
