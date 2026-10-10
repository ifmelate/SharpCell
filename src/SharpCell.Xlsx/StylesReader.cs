using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using SharpCell.Functions;

namespace SharpCell.Xlsx;

/// <summary>A cell style of <c>cellXfs</c>: its number format (null for General) and how it looks.</summary>
internal sealed record CellXf(NumberFormat? Format, CellStyle Style);

/// <summary>
/// The cell styles of <c>xl/styles.xml</c>: a cell's <c>s</c> attribute is an index into
/// <c>cellXfs</c>, whose <c>numFmtId</c> is a custom code from <c>numFmts</c> or one of Excel's
/// built-in formats, and whose <c>fontId</c>, <c>fillId</c> and <c>borderId</c> point into
/// <c>fonts</c>, <c>fills</c> and <c>borders</c>. Colours are turned into RGB with the theme and the
/// indexed palette. Anything missing or unreadable takes its default, so a file is never refused
/// over a style.
/// </summary>
internal static class StylesReader
{
    // Built-in formats as Excel shows them with an English (United States) locale. Ids 14 and 22
    // follow the locale's short date; OOXML lists them as mm-dd-yy, which Excel does not show.
    private static readonly Dictionary<int, string> BuiltIn = new()
    {
        [1] = "0", [2] = "0.00", [3] = "#,##0", [4] = "#,##0.00", [9] = "0%", [10] = "0.00%",
        [11] = "0.00E+00", [12] = "# ?/?", [13] = "# ??/??", [14] = "m/d/yyyy", [15] = "d-mmm-yy",
        [16] = "d-mmm", [17] = "mmm-yy", [18] = "h:mm AM/PM", [19] = "h:mm:ss AM/PM", [20] = "h:mm",
        [21] = "h:mm:ss", [22] = "m/d/yyyy h:mm", [37] = "#,##0 ;(#,##0)", [38] = "#,##0 ;[Red](#,##0)",
        [39] = "#,##0.00;(#,##0.00)", [40] = "#,##0.00;[Red](#,##0.00)", [45] = "mm:ss", [46] = "[h]:mm:ss",
        [47] = "mmss.0", [48] = "##0.0E+0", [49] = "@",
    };

    // The 64 indexed colours of OOXML, used unless the file has indexedColors; 64 and 65 are the
    // system's text and background colours.
    private static readonly int[] DefaultPalette =
    [
        0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
        0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
        0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
        0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
        0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
        0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
        0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
        0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333,
        0x000000, 0xFFFFFF,
    ];

    private static readonly Dictionary<string, CellBorderStyle> BorderStyles = new(StringComparer.Ordinal)
    {
        ["thin"] = CellBorderStyle.Thin, ["medium"] = CellBorderStyle.Medium, ["thick"] = CellBorderStyle.Thick,
        ["double"] = CellBorderStyle.Double, ["hair"] = CellBorderStyle.Hair, ["dotted"] = CellBorderStyle.Dotted,
        ["dashed"] = CellBorderStyle.Dashed, ["dashDot"] = CellBorderStyle.DashDot, ["dashDotDot"] = CellBorderStyle.DashDotDot,
        ["mediumDashed"] = CellBorderStyle.MediumDashed, ["mediumDashDot"] = CellBorderStyle.MediumDashDot,
        ["mediumDashDotDot"] = CellBorderStyle.MediumDashDotDot, ["slantDashDot"] = CellBorderStyle.SlantDashDot,
    };

    private static readonly Dictionary<string, CellHorizontalAlignment> Horizontal = new(StringComparer.Ordinal)
    {
        ["general"] = CellHorizontalAlignment.General, ["left"] = CellHorizontalAlignment.Left,
        ["center"] = CellHorizontalAlignment.Center, ["right"] = CellHorizontalAlignment.Right,
        ["fill"] = CellHorizontalAlignment.Fill, ["justify"] = CellHorizontalAlignment.Justify,
        ["centerContinuous"] = CellHorizontalAlignment.CenterContinuous, ["distributed"] = CellHorizontalAlignment.Distributed,
    };

    private static readonly Dictionary<string, CellVerticalAlignment> Vertical = new(StringComparer.Ordinal)
    {
        ["top"] = CellVerticalAlignment.Top, ["center"] = CellVerticalAlignment.Center, ["bottom"] = CellVerticalAlignment.Bottom,
        ["justify"] = CellVerticalAlignment.Justify, ["distributed"] = CellVerticalAlignment.Distributed,
    };

    private static readonly Dictionary<string, CellUnderline> Underlines = new(StringComparer.Ordinal)
    {
        ["none"] = CellUnderline.None, ["single"] = CellUnderline.Single, ["double"] = CellUnderline.Double,
        ["singleAccounting"] = CellUnderline.SingleAccounting, ["doubleAccounting"] = CellUnderline.DoubleAccounting,
    };

    /// <summary>
    /// Each <c>cellXfs</c> entry, in order. A part that is not valid XML gives no styles: every cell
    /// is General and has the default style.
    /// </summary>
    public static List<CellXf> Read(Package package, string part, CellColor[] theme)
    {
        try
        {
            using var reader = package.OpenXml(part);
            return ReadStyles(XDocument.Load(reader).Root, theme);
        }
        catch (XmlException)
        {
            return [];
        }
    }

    private static List<CellXf> ReadStyles(XElement? root, CellColor[] theme)
    {
        if (root is null)
            return [];

        // Only numFmts defines codes: differential formats (dxfs) carry numFmt elements too.
        var custom = new Dictionary<int, string>();
        foreach (var numFmt in Items(root, "numFmts", "numFmt"))
        {
            if (Int(numFmt, "numFmtId") is { } id && numFmt.Attribute("formatCode")?.Value is { } code)
                custom[id] = code;
        }

        var palette = Items(root, "colors", "indexedColors").FirstOrDefault() is { } indexed
            ? Children(indexed, "rgbColor").Select(c => ParseRgb(c.Attribute("rgb")?.Value) ?? default).ToArray()
            : DefaultPalette.Select(CellColor.FromRgb).ToArray();
        var colors = new Colors(theme, palette);

        var fonts = Items(root, "fonts", "font").Select(colors.Font).ToList();
        var fills = Items(root, "fills", "fill").Select(colors.Fill).ToList();
        var borders = Items(root, "borders", "border").Select(colors.Border).ToList();

        var styles = new List<CellXf>();
        var shared = new Dictionary<CellStyle, CellStyle>();
        foreach (var xf in Items(root, "cellXfs", "xf"))
        {
            // An id the file does not have falls back to its first entry, the Normal style's.
            var font = At(fonts, Int(xf, "fontId") ?? 0) ?? At(fonts, 0) ?? CellFont.Default;
            var border = At(borders, Int(xf, "borderId") ?? 0) ?? At(borders, 0) ?? Edges.None;
            var style = CellStyle.Default with
            {
                Font = font,
                Fill = (At(fills, Int(xf, "fillId") ?? 0) ?? At(fills, 0))?.Color,
                LeftBorder = border.Left,
                RightBorder = border.Right,
                TopBorder = border.Top,
                BottomBorder = border.Bottom,
            };
            if (Child(xf, "alignment") is { } alignment)
            {
                style = style with
                {
                    HorizontalAlignment = Lookup(Horizontal, alignment.Attribute("horizontal")?.Value) ?? CellHorizontalAlignment.General,
                    VerticalAlignment = Lookup(Vertical, alignment.Attribute("vertical")?.Value) ?? CellVerticalAlignment.Bottom,
                    WrapText = IsTrue(alignment.Attribute("wrapText")?.Value),
                    Indent = Int(alignment, "indent") is { } indent and <= 250 ? indent : 0,
                };
            }

            // Equal styles share one instance: many cellXfs differ only in their number format.
            if (shared.TryGetValue(style, out var same))
                style = same;
            else
                shared.Add(style, style);
            styles.Add(new CellXf(Format(Int(xf, "numFmtId") ?? 0, custom), style));
        }

        return styles;
    }

    /// <summary>An <c>RRGGBB</c> or <c>AARRGGBB</c> colour; the alpha is ignored, as Excel does.</summary>
    public static CellColor? ParseRgb(string? text)
    {
        if (text is null || (text.Length != 6 && text.Length != 8))
            return null;
        return int.TryParse(text.AsSpan(text.Length - 6), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var rgb)
            ? CellColor.FromRgb(rgb)
            : null;
    }

    /// <summary>
    /// A colour with its lightness moved by a tint from -1 (black) to 1 (white), the way OOXML
    /// defines it: the colour in hue, lightness and saturation, the lightness scaled towards 0 or 1.
    /// </summary>
    public static CellColor Tint(CellColor color, double tint)
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

    private static IEnumerable<XElement> Items(XElement root, string section, string item) =>
        Children(root, section).Take(1).SelectMany(s => Children(s, item));

    private static IEnumerable<XElement> Children(XElement parent, string localName) =>
        parent.Elements().Where(e => e.Name.LocalName == localName);

    private static XElement? Child(XElement parent, string localName) => Children(parent, localName).FirstOrDefault();

    private static T? At<T>(List<T> list, int index) where T : class => index >= 0 && index < list.Count ? list[index] : null;

    private static int? Int(XElement element, string attribute) =>
        int.TryParse(element.Attribute(attribute)?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value) ? value : null;

    private static TValue? Lookup<TValue>(Dictionary<string, TValue> map, string? key) where TValue : struct =>
        key is not null && map.TryGetValue(key, out var value) ? value : null;

    private static bool IsTrue(string? value) => value is "1" or "true";

    // An effect such as <b/> is on unless its val says otherwise.
    private static bool Flag(XElement font, string localName) =>
        Child(font, localName) is { } flag && (flag.Attribute("val")?.Value is not { } value || IsTrue(value));

    private static NumberFormat? Format(int id, Dictionary<int, string> custom)
    {
        if (!custom.TryGetValue(id, out var code) && !BuiltIn.TryGetValue(id, out code))
            return null;
        if (code.Equals("General", StringComparison.OrdinalIgnoreCase))
            return null;
        return NumberFormat.TryParse(code, out var format) ? format : null;
    }

    private sealed record FillColor(CellColor? Color);

    private sealed record Edges(CellBorder? Left, CellBorder? Right, CellBorder? Top, CellBorder? Bottom)
    {
        public static readonly Edges None = new(null, null, null, null);
    }

    /// <summary>Turns the colour elements of fonts, fills and borders into RGB.</summary>
    private sealed class Colors(CellColor[] theme, CellColor[] palette)
    {
        public CellFont Font(XElement font)
        {
            var result = CellFont.Default with
            {
                Bold = Flag(font, "b"),
                Italic = Flag(font, "i"),
                Strikethrough = Flag(font, "strike"),
                Underline = Child(font, "u") is { } u ? Lookup(Underlines, u.Attribute("val")?.Value ?? "single") ?? CellUnderline.Single : CellUnderline.None,
                Color = Color(Child(font, "color")),
            };
            if (Child(font, "name")?.Attribute("val")?.Value is { } name && !string.IsNullOrWhiteSpace(name))
                result = result with { Name = name };
            if (double.TryParse(Child(font, "sz")?.Attribute("val")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var size)
                && size is >= 1 and <= 409)
                result = result with { Size = size };
            return result;
        }

        // Only solid fills: patterns and gradients have no single colour to show.
        public FillColor Fill(XElement fill) =>
            Child(fill, "patternFill") is { } pattern && pattern.Attribute("patternType")?.Value == "solid"
                ? new FillColor(Color(Child(pattern, "fgColor")))
                : new FillColor(null);

        public Edges Border(XElement border) => new(
            Edge(Child(border, "left") ?? Child(border, "start")),
            Edge(Child(border, "right") ?? Child(border, "end")),
            Edge(Child(border, "top")),
            Edge(Child(border, "bottom")));

        private CellBorder? Edge(XElement? edge) =>
            edge is not null && Lookup(BorderStyles, edge.Attribute("style")?.Value) is { } style
                ? new CellBorder(style, Color(Child(edge, "color")))
                : null;

        private CellColor? Color(XElement? color)
        {
            if (color is null || IsTrue(color.Attribute("auto")?.Value))
                return null;

            CellColor? found;
            if (color.Attribute("rgb") is { } rgb)
                found = ParseRgb(rgb.Value);
            else if (Int(color, "theme") is { } index)
                found = index < theme.Length ? theme[index] : null;
            else if (Int(color, "indexed") is { } entry)
                found = entry < palette.Length ? palette[entry] : entry < DefaultPalette.Length ? CellColor.FromRgb(DefaultPalette[entry]) : null;
            else
                found = null;

            if (found is { } value
                && double.TryParse(color.Attribute("tint")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var tint)
                && tint is >= -1 and <= 1 and not 0)
                return Tint(value, tint);
            return found;
        }
    }
}
