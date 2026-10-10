using System.Globalization;

namespace SharpCell;

/// <summary>
/// A colour as Excel shows it, in sRGB. Theme and indexed colours of a file are turned into
/// these when it is read; Excel ignores the alpha channel, so there is none.
/// </summary>
/// <param name="R">Red, 0 to 255.</param>
/// <param name="G">Green, 0 to 255.</param>
/// <param name="B">Blue, 0 to 255.</param>
public readonly record struct CellColor(byte R, byte G, byte B)
{
    /// <summary>A colour from a <c>0xRRGGBB</c> number.</summary>
    public static CellColor FromRgb(int rgb) => new((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb);

    /// <summary>The colour as <c>#RRGGBB</c>.</summary>
    public override string ToString() => string.Create(CultureInfo.InvariantCulture, $"#{R:X2}{G:X2}{B:X2}");
}
