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
/// A fill as the model holds it: a colour (solid, or the background of a pattern), a pattern and its
/// colour, or a gradient. <see cref="Other"/> marks a fill the model cannot hold, which is never reused.
/// </summary>
internal sealed record FillLook(CellColor? Color, CellFillPattern Pattern, CellColor? PatternColor, CellGradient? Gradient, bool Other = false)
{
    public static readonly FillLook None = new(null, CellFillPattern.Solid, null, null);

    public static FillLook Of(CellStyle style) => new(style.Fill, style.FillPattern, style.PatternColor, style.Gradient);
}

/// <summary>The four edges of a border as read.</summary>
internal sealed record BorderEdges(CellBorder? Left, CellBorder? Right, CellBorder? Top, CellBorder? Bottom)
{
    public static readonly BorderEdges None = new(null, null, null, null);
}

/// <summary>
/// The parts of <c>styles.xml</c> as read, by index: custom number formats, fonts, fills (the colour
/// of a solid fill, and whether a fill is empty), borders and cell styles.
/// </summary>
internal sealed class StyleSheet
{
    public Dictionary<int, string> Custom { get; } = [];

    public List<CellFont> Fonts { get; } = [];

    /// <summary>The fills as the model holds them.</summary>
    public List<FillLook> Fills { get; } = [];

    public List<BorderEdges> Borders { get; } = [];

    public List<CellXf> Xfs { get; } = [];

    /// <summary>The differential formats conditional formats point at, by index.</summary>
    public List<ConditionalStyle> Dxfs { get; } = [];

    /// <summary>The named styles other than Normal, each with its name, in the file's order.</summary>
    public List<CellStyle> NamedStyles { get; } = [];

    /// <summary>Turns a font element (or a rich text run's rPr) into a font.</summary>
    public Func<XElement, CellFont>? Font { get; set; }

    /// <summary>Turns a colour element (rgb, theme or indexed, with a tint) into RGB; null for automatic or unreadable.</summary>
    public Func<XElement?, CellColor?>? Color { get; set; }
}

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
    internal static readonly Dictionary<int, string> BuiltIn = new()
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

    internal static readonly Dictionary<string, CellBorderStyle> BorderStyles = new(StringComparer.Ordinal)
    {
        ["thin"] = CellBorderStyle.Thin, ["medium"] = CellBorderStyle.Medium, ["thick"] = CellBorderStyle.Thick,
        ["double"] = CellBorderStyle.Double, ["hair"] = CellBorderStyle.Hair, ["dotted"] = CellBorderStyle.Dotted,
        ["dashed"] = CellBorderStyle.Dashed, ["dashDot"] = CellBorderStyle.DashDot, ["dashDotDot"] = CellBorderStyle.DashDotDot,
        ["mediumDashed"] = CellBorderStyle.MediumDashed, ["mediumDashDot"] = CellBorderStyle.MediumDashDot,
        ["mediumDashDotDot"] = CellBorderStyle.MediumDashDotDot, ["slantDashDot"] = CellBorderStyle.SlantDashDot,
    };

    internal static readonly Dictionary<string, CellHorizontalAlignment> Horizontal = new(StringComparer.Ordinal)
    {
        ["general"] = CellHorizontalAlignment.General, ["left"] = CellHorizontalAlignment.Left,
        ["center"] = CellHorizontalAlignment.Center, ["right"] = CellHorizontalAlignment.Right,
        ["fill"] = CellHorizontalAlignment.Fill, ["justify"] = CellHorizontalAlignment.Justify,
        ["centerContinuous"] = CellHorizontalAlignment.CenterContinuous, ["distributed"] = CellHorizontalAlignment.Distributed,
    };

    internal static readonly Dictionary<string, CellVerticalAlignment> Vertical = new(StringComparer.Ordinal)
    {
        ["top"] = CellVerticalAlignment.Top, ["center"] = CellVerticalAlignment.Center, ["bottom"] = CellVerticalAlignment.Bottom,
        ["justify"] = CellVerticalAlignment.Justify, ["distributed"] = CellVerticalAlignment.Distributed,
    };

    internal static readonly Dictionary<string, CellUnderline> Underlines = new(StringComparer.Ordinal)
    {
        ["none"] = CellUnderline.None, ["single"] = CellUnderline.Single, ["double"] = CellUnderline.Double,
        ["singleAccounting"] = CellUnderline.SingleAccounting, ["doubleAccounting"] = CellUnderline.DoubleAccounting,
    };

    /// <summary>
    /// Each <c>cellXfs</c> entry, in order. A part that is not valid XML gives no styles: every cell
    /// is General and has the default style.
    /// </summary>
    public static StyleSheet Read(Package package, string part, CellColor[] theme)
    {
        try
        {
            using var reader = package.OpenXml(part);
            return ReadStyleSheet(XDocument.Load(reader).Root, theme);
        }
        catch (XmlException)
        {
            return new StyleSheet();
        }
    }

    /// <summary>Everything the styles part defines, by index; empty for a part without a root.</summary>
    public static StyleSheet ReadStyleSheet(XElement? root, CellColor[] theme)
    {
        var sheet = new StyleSheet();
        if (root is null)
            return sheet;

        // Only numFmts defines codes: differential formats (dxfs) carry numFmt elements too.
        var custom = sheet.Custom;
        foreach (var numFmt in Items(root, "numFmts", "numFmt"))
        {
            if (Int(numFmt, "numFmtId") is { } id && numFmt.Attribute("formatCode")?.Value is { } code)
                custom[id] = code;
        }

        var palette = Items(root, "colors", "indexedColors").FirstOrDefault() is { } indexed
            ? Children(indexed, "rgbColor").Select(c => ParseRgb(c.Attribute("rgb")?.Value) ?? default).ToArray()
            : DefaultPalette.Select(CellColor.FromRgb).ToArray();
        var colors = new Colors(theme, palette);
        sheet.Color = colors.Color;
        sheet.Font = colors.Font;

        var fonts = sheet.Fonts;
        fonts.AddRange(Items(root, "fonts", "font").Select(colors.Font));
        var fills = sheet.Fills;
        fills.AddRange(Items(root, "fills", "fill").Select(colors.Fill));
        var borders = sheet.Borders;
        borders.AddRange(Items(root, "borders", "border").Select(colors.Border));

        // Named styles: cellStyles name entries of cellStyleXfs; a cell style's xfId says which it is based on.
        var masters = Items(root, "cellStyleXfs", "xf").Select(xf => Look(xf, fonts, fills, borders)).ToList();
        var names = new Dictionary<int, string>();
        foreach (var entry in Items(root, "cellStyles", "cellStyle"))
        {
            if (entry.Attribute("name")?.Value is { } name && !string.IsNullOrWhiteSpace(name) && Int(entry, "xfId") is { } id
                && Int(entry, "builtinId") != 0 && !name.Equals("Normal", StringComparison.OrdinalIgnoreCase) && names.TryAdd(id, name)
                && id < masters.Count)
                sheet.NamedStyles.Add(masters[id] with { Name = name });
        }

        var styles = sheet.Xfs;
        var shared = new Dictionary<CellStyle, CellStyle>();
        foreach (var xf in Items(root, "cellXfs", "xf"))
        {
            var style = Look(xf, fonts, fills, borders);
            if (Int(xf, "xfId") is { } master && names.TryGetValue(master, out var styleName))
                style = style with { Name = styleName };

            // Equal styles share one instance: many cellXfs differ only in their number format.
            if (shared.TryGetValue(style, out var same))
                style = same;
            else
                shared.Add(style, style);
            styles.Add(new CellXf(Format(Int(xf, "numFmtId") ?? 0, custom), style));
        }

        sheet.Dxfs.AddRange(Items(root, "dxfs", "dxf").Select(colors.Dxf));

        return sheet;
    }

    private static CellStyle Look(XElement xf, List<CellFont> fonts, List<FillLook> fills, List<BorderEdges> borders)
    {
        // An id the file does not have falls back to its first entry, the Normal style's.
        var font = At(fonts, Int(xf, "fontId") ?? 0) ?? At(fonts, 0) ?? CellFont.Default;
        var border = At(borders, Int(xf, "borderId") ?? 0) ?? At(borders, 0) ?? BorderEdges.None;
        var fill = At(fills, Int(xf, "fillId") ?? 0) ?? At(fills, 0) ?? FillLook.None;
        var style = CellStyle.Default with
        {
            Font = font,
            Fill = fill.Color,
            FillPattern = fill.Pattern,
            PatternColor = fill.PatternColor,
            Gradient = fill.Gradient,
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
                ShrinkToFit = IsTrue(alignment.Attribute("shrinkToFit")?.Value),
            };
            // 0 to 90 turn the text up, 91 to 180 down by the amount past 90, 255 stacks it.
            switch (Int(alignment, "textRotation"))
            {
                case 255:
                    style = style with { VerticalText = true };
                    break;
                case { } up and <= 90:
                    style = style with { TextRotation = up };
                    break;
                case { } down and <= 180:
                    style = style with { TextRotation = 90 - down };
                    break;
            }
        }
        if (Child(xf, "protection") is { } protection)
        {
            style = style with
            {
                Locked = protection.Attribute("locked") is null || IsTrue(protection.Attribute("locked")!.Value),
                FormulaHidden = IsTrue(protection.Attribute("hidden")?.Value),
            };
        }

        return style;
    }

    /// <summary>Colours for a file without a styles part: the theme's and the default indexed palette.</summary>
    public static Func<XElement?, CellColor?> DefaultColors(CellColor[] theme) => new Colors(theme, DefaultPalette.Select(CellColor.FromRgb).ToArray()).Color;

    /// <summary>Fonts for a file without a styles part: colours from the theme and the default indexed palette.</summary>
    public static Func<XElement, CellFont> DefaultFonts(CellColor[] theme) => new Colors(theme, DefaultPalette.Select(CellColor.FromRgb).ToArray()).Font;

    /// <summary>An <c>RRGGBB</c> or <c>AARRGGBB</c> colour; the alpha is ignored, as Excel does.</summary>
    public static CellColor? ParseRgb(string? text)
    {
        if (text is null || (text.Length != 6 && text.Length != 8))
            return null;
        return int.TryParse(text.AsSpan(text.Length - 6), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var rgb)
            ? CellColor.FromRgb(rgb)
            : null;
    }

    /// <summary>A colour with its lightness moved by a tint from -1 (black) to 1 (white), as OOXML defines it.</summary>
    public static CellColor Tint(CellColor color, double tint) => CellColor.Tinted(color, tint);

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

    internal static readonly Dictionary<string, CellFillPattern> Patterns = new(StringComparer.Ordinal)
    {
        ["solid"] = CellFillPattern.Solid, ["darkGray"] = CellFillPattern.DarkGray, ["mediumGray"] = CellFillPattern.MediumGray,
        ["lightGray"] = CellFillPattern.LightGray, ["gray125"] = CellFillPattern.Gray125, ["gray0625"] = CellFillPattern.Gray0625,
        ["darkHorizontal"] = CellFillPattern.DarkHorizontal, ["darkVertical"] = CellFillPattern.DarkVertical,
        ["darkDown"] = CellFillPattern.DarkDown, ["darkUp"] = CellFillPattern.DarkUp, ["darkGrid"] = CellFillPattern.DarkGrid,
        ["darkTrellis"] = CellFillPattern.DarkTrellis, ["lightHorizontal"] = CellFillPattern.LightHorizontal,
        ["lightVertical"] = CellFillPattern.LightVertical, ["lightDown"] = CellFillPattern.LightDown, ["lightUp"] = CellFillPattern.LightUp,
        ["lightGrid"] = CellFillPattern.LightGrid, ["lightTrellis"] = CellFillPattern.LightTrellis,
    };

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
                Script = Child(font, "vertAlign")?.Attribute("val")?.Value switch
                {
                    "superscript" => CellFontScript.Superscript,
                    "subscript" => CellFontScript.Subscript,
                    _ => CellFontScript.Normal,
                },
            };
            if ((Child(font, "name") ?? Child(font, "rFont"))?.Attribute("val")?.Value is { } name && !string.IsNullOrWhiteSpace(name))
                result = result with { Name = name };
            if (double.TryParse(Child(font, "sz")?.Attribute("val")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var size)
                && size is >= 1 and <= 409)
                result = result with { Size = size };
            return result;
        }

        // A differential format: only what it sets. Its fill colour is the pattern's background, which
        // Excel shows for a solid or unstated pattern; <b val="0"/> turns bold off.
        public ConditionalStyle Dxf(XElement dxf)
        {
            var font = Child(dxf, "font");
            var pattern = Child(dxf, "fill") is { } fill ? Child(fill, "patternFill") : null;
            var patternType = pattern?.Attribute("patternType")?.Value;
            var border = Child(dxf, "border") is { } edges ? Border(edges) : BorderEdges.None;
            return new ConditionalStyle
            {
                Bold = font is null ? null : State(font, "b"),
                Italic = font is null ? null : State(font, "i"),
                Strikethrough = font is null ? null : State(font, "strike"),
                Underline = font is not null && Child(font, "u") is { } u ? Lookup(Underlines, u.Attribute("val")?.Value ?? "single") ?? CellUnderline.Single : null,
                FontColor = font is null ? null : Color(Child(font, "color")),
                Fill = pattern is not null && patternType is null or "solid" ? Color(Child(pattern, "bgColor") ?? Child(pattern, "fgColor")) : null,
                LeftBorder = border.Left,
                RightBorder = border.Right,
                TopBorder = border.Top,
                BottomBorder = border.Bottom,
                NumberFormat = Child(dxf, "numFmt")?.Attribute("formatCode")?.Value,
            };

            static bool? State(XElement font, string name) => Child(font, name) is { } flag ? flag.Attribute("val") is not { } val || IsTrue(val.Value) : null;
        }

        // A solid fill's colour is its fgColor; a pattern is drawn in fgColor over bgColor. Indexed 64
        // and 65 are the system's automatic colours, which the model holds as null.
        public FillLook Fill(XElement fill)
        {
            if (Child(fill, "gradientFill") is { } gradient)
                return Gradient(gradient) is { } read ? new FillLook(null, CellFillPattern.Solid, null, read) : FillLook.None with { Other = true };
            if (Child(fill, "patternFill") is not { } pattern)
                return FillLook.None with { Other = true };
            var type = pattern.Attribute("patternType")?.Value ?? "none";
            if (type == "none")
                return FillLook.None;
            if (Lookup(Patterns, type) is not { } kind)
                return FillLook.None with { Other = true };
            return kind == CellFillPattern.Solid
                ? new FillLook(Automatic(Child(pattern, "fgColor")), kind, null, null)
                : new FillLook(Automatic(Child(pattern, "bgColor")), kind, Automatic(Child(pattern, "fgColor")), null);
        }

        private CellColor? Automatic(XElement? color) => Int(color ?? new XElement("x"), "indexed") is 64 or 65 ? null : Color(color);

        private CellGradient? Gradient(XElement element)
        {
            var stops = new List<GradientStop>();
            foreach (var stop in Children(element, "stop"))
            {
                if (!double.TryParse(stop.Attribute("position")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var position) || position is < 0 or > 1)
                    return null;
                stops.Add(new GradientStop(position, Color(Child(stop, "color")) ?? new CellColor(0, 0, 0)));
            }

            if (stops.Count < 2)
                return null;
            return new CellGradient(stops)
            {
                Type = element.Attribute("type")?.Value == "path" ? CellGradientType.Path : CellGradientType.Linear,
                Angle = Number(element, "degree"),
                Left = Edge(element, "left"),
                Right = Edge(element, "right"),
                Top = Edge(element, "top"),
                Bottom = Edge(element, "bottom"),
            };

            static double Number(XElement element, string name) =>
                double.TryParse(element.Attribute(name)?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) && double.IsFinite(value) ? value : 0;

            static double Edge(XElement element, string name) => Math.Clamp(Number(element, name), 0, 1);
        }

        public BorderEdges Border(XElement border) => new(
            Edge(Child(border, "left") ?? Child(border, "start")),
            Edge(Child(border, "right") ?? Child(border, "end")),
            Edge(Child(border, "top")),
            Edge(Child(border, "bottom")));

        private CellBorder? Edge(XElement? edge) =>
            edge is not null && Lookup(BorderStyles, edge.Attribute("style")?.Value) is { } style
                ? new CellBorder(style, Color(Child(edge, "color")))
                : null;

        public CellColor? Color(XElement? color)
        {
            if (color is null || IsTrue(color.Attribute("auto")?.Value))
                return null;

            CellColor? found;
            var tinted = double.TryParse(color.Attribute("tint")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var tint)
                && tint is >= -1 and <= 1 and not 0;
            if (color.Attribute("rgb") is { } rgb)
                found = ParseRgb(rgb.Value);
            else if (Int(color, "theme") is { } index)
                return index < theme.Length ? CellColor.OfTheme(theme[index], (ThemeColor)index, tinted ? tint : 0) : null;
            else if (Int(color, "indexed") is { } entry)
                found = entry < palette.Length ? palette[entry] : entry < DefaultPalette.Length ? CellColor.FromRgb(DefaultPalette[entry]) : null;
            else
                found = null;

            if (found is { } value && tinted)
                return Tint(value, tint);
            return found;
        }
    }
}
