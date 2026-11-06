using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using SharpCell.Functions;

namespace SharpCell.Xlsx;

/// <summary>
/// Finds or adds the cell style (<c>cellXfs</c> entry) a cell needs. A new style starts from the
/// cell's old one and replaces only the font, fill, border, alignment, protection or number format that
/// differ, so what SharpCell does not model (pattern fills, rotated text) is kept.
/// Fonts, fills, borders, number formats and styles equal to existing ones are reused. The part is
/// read only when a style changes, and created when the file has none.
/// </summary>
internal sealed class StylesWriter
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

    // CT_Stylesheet's children in schema order.
    private static readonly string[] Order =
        ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"];

    private const string Empty =
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        + "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>"
        + "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
        + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
        + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
        + "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>"
        + "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
        + "</styleSheet>";

    private readonly PackageEdit _edit;
    private readonly string _workbookPart;
    private readonly CellColor[] _theme;

    // Without a theme part a theme colour cannot be referred to: it is saved as its RGB.
    private readonly bool _hasTheme;
    private string? _part;
    private XDocument? _document;
    private StyleSheet _read = new();
    private List<XElement> _xfs = [];
    private List<XElement> _masters = [];
    private Dictionary<string, int> _named = new(StringComparer.OrdinalIgnoreCase);
    private bool _changed;
    private readonly Dictionary<(int, CellStyle, string), int> _cache = [];

    // Excel's built-in cell styles by name, with the builtinId that makes Excel treat them as its own.
    private static readonly Dictionary<string, int> BuiltIn = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Comma"] = 3, ["Currency"] = 4, ["Percent"] = 5, ["Comma [0]"] = 6, ["Currency [0]"] = 7, ["Hyperlink"] = 8,
        ["Followed Hyperlink"] = 9, ["Note"] = 10, ["Warning Text"] = 11, ["Title"] = 15, ["Heading 1"] = 16, ["Heading 2"] = 17,
        ["Heading 3"] = 18, ["Heading 4"] = 19, ["Input"] = 20, ["Output"] = 21, ["Calculation"] = 22, ["Check Cell"] = 23,
        ["Linked Cell"] = 24, ["Total"] = 25, ["Good"] = 26, ["Bad"] = 27, ["Neutral"] = 28, ["Accent1"] = 29, ["20% - Accent1"] = 30,
        ["40% - Accent1"] = 31, ["60% - Accent1"] = 32, ["Accent2"] = 33, ["20% - Accent2"] = 34, ["40% - Accent2"] = 35,
        ["60% - Accent2"] = 36, ["Accent3"] = 37, ["20% - Accent3"] = 38, ["40% - Accent3"] = 39, ["60% - Accent3"] = 40,
        ["Accent4"] = 41, ["20% - Accent4"] = 42, ["40% - Accent4"] = 43, ["60% - Accent4"] = 44, ["Accent5"] = 45,
        ["20% - Accent5"] = 46, ["40% - Accent5"] = 47, ["60% - Accent5"] = 48, ["Accent6"] = 49, ["20% - Accent6"] = 50,
        ["40% - Accent6"] = 51, ["60% - Accent6"] = 52, ["Explanatory Text"] = 53,
    };

    private StylesWriter(PackageEdit edit, string workbookPart, string? part, CellColor[] theme, bool hasTheme)
    {
        _hasTheme = hasTheme;
        _edit = edit;
        _workbookPart = workbookPart;
        _part = part;
        _theme = theme;
    }

    public static StylesWriter Open(PackageEdit edit, string workbookPart, Dictionary<string, Relationship> relationships)
    {
        var part = WorkbookReader.FirstOfType(relationships, "styles")?.Target;
        var themePart = WorkbookReader.FirstOfType(relationships, "theme")?.Target;
        var theme = ThemeReader.Read(edit.Package, themePart);
        return new StylesWriter(edit, workbookPart, part is not null && edit.Package.Exists(part) ? part : null, theme,
            themePart is not null && edit.Package.Exists(themePart));
    }

    /// <summary>The index of a cell style with this look and format, based on the style at <paramref name="baseIndex"/>.</summary>
    /// <param name="baseIndex">The cell's style in the file, or 0.</param>
    /// <param name="style">How the cell looks.</param>
    /// <param name="formatCode">The number format code; <c>General</c> for none.</param>
    public int XfFor(int baseIndex, CellStyle style, string formatCode)
    {
        if (_cache.TryGetValue((baseIndex, style, formatCode), out var cached))
            return cached;
        var key = style;
        style = Portable(style);
        Load();
        if (baseIndex < 0 || baseIndex >= _xfs.Count)
            baseIndex = 0;

        var xf = new XElement(_xfs[baseIndex]);
        Apply(xf, _read.Xfs[baseIndex], style, formatCode);
        var index = _xfs.FindIndex(existing => XNode.DeepEquals(existing, xf));
        if (index < 0)
        {
            index = _xfs.Count;
            Section("cellXfs").Add(xf);
            _xfs.Add(xf);
            _read.Xfs.Add(new CellXf(ParseFormat(formatCode), style));
            _changed = true;
        }

        _cache[(baseIndex, key, formatCode)] = index;
        return index;
    }

    /// <summary>The differential format at an index of the file, or null when it has none there.</summary>
    public ConditionalStyle? DxfAt(int index)
    {
        Load();
        return index >= 0 && index < _read.Dxfs.Count ? _read.Dxfs[index] : null;
    }

    /// <summary>The index of a differential format for a conditional format's look, added when the file has none like it.</summary>
    public int DxfFor(ConditionalStyle style)
    {
        if (!_hasTheme)
        {
            style = style with
            {
                FontColor = Plain(style.FontColor), Fill = Plain(style.Fill), LeftBorder = Plain(style.LeftBorder),
                RightBorder = Plain(style.RightBorder), TopBorder = Plain(style.TopBorder), BottomBorder = Plain(style.BottomBorder),
            };
        }

        Load();
        var index = _read.Dxfs.IndexOf(style);
        if (index >= 0)
            return index;

        // The schema's order: font, numFmt, fill, border; in the font b, i, strike, u, color.
        XElement? font = null;
        if (style.Bold is not null || style.Italic is not null || style.Strikethrough is not null || style.Underline is not null || style.FontColor is not null)
        {
            font = new XElement(Main + "font",
                Flag("b", style.Bold), Flag("i", style.Italic), Flag("strike", style.Strikethrough),
                style.Underline is { } underline
                    ? new XElement(Main + "u", underline == CellUnderline.Single ? null : new XAttribute("val", underline == CellUnderline.None ? "none" : Name(StylesReader.Underlines, underline)))
                    : null,
                style.FontColor is { } color ? Color("color", color) : null);
        }

        var element = new XElement(Main + "dxf", font,
            style.NumberFormat is { } code ? new XElement(Main + "numFmt", new XAttribute("numFmtId", NumberFormatId(code)), new XAttribute("formatCode", code)) : null,
            style.Fill is { } fill ? new XElement(Main + "fill", new XElement(Main + "patternFill", Color("bgColor", fill))) : null,
            (style.LeftBorder ?? style.RightBorder ?? style.TopBorder ?? style.BottomBorder) is null
                ? null
                : new XElement(Main + "border", DxfEdge("left", style.LeftBorder), DxfEdge("right", style.RightBorder), DxfEdge("top", style.TopBorder), DxfEdge("bottom", style.BottomBorder)));
        return Add("dxfs", element, _read.Dxfs, style);

        static XElement? Flag(string name, bool? on) => on is { } value ? new XElement(Main + name, value ? null : new XAttribute("val", "0")) : null;

        static XElement? DxfEdge(string name, CellBorder? border) => border is null ? null : Edge(name, border);
    }

    /// <summary>Changes the first cell style, which cells without a style of their own show, and the Normal style.</summary>
    public void SetDefault(CellStyle style)
    {
        style = Portable(style);
        Load();
        var old = _read.Xfs[0];
        Apply(_xfs[0], old, style, old.Format?.FormatCode ?? "General");
        if (Section("cellStyleXfs").Elements(Main + "xf").FirstOrDefault() is { } normal)
            Apply(normal, old, style, old.Format?.FormatCode ?? "General");
        _read.Xfs[0] = old with { Style = style };
        _cache.Clear();
        _changed = true;
    }

    /// <summary>Puts the changed or created part into the edit.</summary>
    public void Save()
    {
        if (!_changed)
            return;
        foreach (var section in _document!.Root!.Elements())
        {
            if (section.Attribute("count") is not null || section.Name.LocalName is "numFmts" or "fonts" or "fills" or "borders" or "cellStyleXfs" or "cellXfs" or "cellStyles" or "dxfs")
                section.SetAttributeValue("count", section.Elements().Count());
        }

        if (_part is null)
        {
            _part = _workbookPart[..(_workbookPart.LastIndexOf('/') + 1)] + "styles.xml";
            _edit.AddPart(_workbookPart, _part, XmlParts.Save(_document), PackageEdit.OfficeRelationships + "styles", ContentType);
        }
        else
        {
            _edit.Parts[_part] = XmlParts.Save(_document);
        }
    }

    private void Load()
    {
        if (_document is not null)
            return;
        _document = (_part is null ? null : _edit.ReadXml(_part)) ?? XDocument.Parse(Empty);
        if (_document.Root is null || _document.Root.Name.Namespace != Main)
            throw new System.IO.InvalidDataException("The styles part is not a SpreadsheetML style sheet.");
        _read = StylesReader.ReadStyleSheet(_document.Root, _theme);
        _xfs = Section("cellXfs").Elements(Main + "xf").ToList();
        _masters = _document.Root.Element(Main + "cellStyleXfs")?.Elements(Main + "xf").ToList() ?? [];
        foreach (var entry in _document.Root.Element(Main + "cellStyles")?.Elements(Main + "cellStyle") ?? [])
        {
            if ((string?)entry.Attribute("name") is { } name && int.TryParse((string?)entry.Attribute("xfId"), NumberStyles.None, CultureInfo.InvariantCulture, out var id))
                _named.TryAdd(name, id);
        }
        if (_xfs.Count == 0)
        {
            var xf = new XElement(Main + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", Font(CellFont.Default)),
                new XAttribute("fillId", Fill(FillLook.None)), new XAttribute("borderId", Border(BorderEdges.None)));
            Section("cellXfs").Add(xf);
            _xfs.Add(xf);
            _read.Xfs.Add(new CellXf(null, CellStyle.Default with { Font = CellFont.Default }));
            _changed = true;
        }
    }

    // Replaces in the style what differs from what it was read as.
    private void Apply(XElement xf, CellXf old, CellStyle style, string formatCode)
    {
        var was = old.Style;
        if (style.Font != was.Font)
            Set(xf, "fontId", Font(style.Font), "applyFont");
        if (FillLook.Of(style) != FillLook.Of(was))
            Set(xf, "fillId", Fill(FillLook.Of(style)), "applyFill");
        var edges = new BorderEdges(style.LeftBorder, style.RightBorder, style.TopBorder, style.BottomBorder);
        if (edges != new BorderEdges(was.LeftBorder, was.RightBorder, was.TopBorder, was.BottomBorder))
            Set(xf, "borderId", Border(edges), "applyBorder");
        if ((style.HorizontalAlignment, style.VerticalAlignment, style.WrapText, style.Indent, style.TextRotation, style.VerticalText, style.ShrinkToFit)
            != (was.HorizontalAlignment, was.VerticalAlignment, was.WrapText, was.Indent, was.TextRotation, was.VerticalText, was.ShrinkToFit))
            SetAlignment(xf, style);
        if ((style.Locked, style.FormulaHidden) != (was.Locked, was.FormulaHidden))
            SetProtection(xf, style);
        if (formatCode != (old.Format?.FormatCode ?? "General"))
            Set(xf, "numFmtId", NumberFormatId(formatCode), "applyNumberFormat");
        if (!string.Equals(style.Name, was.Name, StringComparison.OrdinalIgnoreCase))
            xf.SetAttributeValue("xfId", style.Name is { } name ? MasterFor(name, style) : 0);
    }

    /// <summary>
    /// Adds the workbook's named styles the file lacks, changes the look of those whose look changed,
    /// and drops those removed from the workbook (their cell styles are no longer used).
    /// </summary>
    public void SetNamedStyles(IReadOnlyDictionary<string, CellStyle> now, IReadOnlyDictionary<string, CellStyle> loaded)
    {
        if (now.Count == loaded.Count && now.All(pair => loaded.TryGetValue(pair.Key, out var was) && was == pair.Value))
            return;
        Load();
        foreach (var (name, named) in now)
        {
            var style = Portable(named);
            if (!_named.TryGetValue(name, out var index))
            {
                MasterFor(name, style);
            }
            else if (loaded.TryGetValue(name, out var was) && was != named && index < _masters.Count)
            {
                Apply(_masters[index], new CellXf(null, was), style, "General");
                foreach (var entry in _document!.Root!.Element(Main + "cellStyles")?.Elements(Main + "cellStyle") ?? [])
                {
                    if (string.Equals((string?)entry.Attribute("name"), name, StringComparison.OrdinalIgnoreCase))
                        entry.SetAttributeValue("name", name);   // the spelling the workbook has now
                }

                _cache.Clear();
                _changed = true;
            }
        }

        foreach (var name in loaded.Keys.Where(name => !now.ContainsKey(name)))
        {
            _document!.Root!.Element(Main + "cellStyles")?.Elements(Main + "cellStyle")
                .Where(entry => string.Equals((string?)entry.Attribute("name"), name, StringComparison.OrdinalIgnoreCase)).Remove();
            _named.Remove(name);
            _changed = true;
        }
    }

    // The cellStyleXfs entry of a named style, added with its cellStyles entry when the file has none.
    private int MasterFor(string name, CellStyle look)
    {
        if (_named.TryGetValue(name, out var existing))
            return existing;
        var section = Section("cellStyleXfs");
        if (_masters.Count == 0)
        {
            var normal = new XElement(Main + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0));
            section.Add(normal);
            _masters.Add(normal);
        }

        var master = new XElement(_masters[0]);
        master.Attributes().Where(a => a.Name.LocalName.StartsWith("apply", StringComparison.Ordinal)).Remove();
        var basis = _read.Xfs.Count > 0 ? _read.Xfs[0] with { Format = null } : new CellXf(null, CellStyle.Default);
        Apply(master, basis, look with { Name = basis.Style.Name }, "General");   // the look only: a style entry has no xfId
        section.Add(master);
        _masters.Add(master);
        var index = _masters.Count - 1;
        var styles = Section("cellStyles");
        if (!styles.Elements(Main + "cellStyle").Any())
            styles.Add(new XElement(Main + "cellStyle", new XAttribute("name", "Normal"), new XAttribute("xfId", 0), new XAttribute("builtinId", 0)));
        styles.Add(new XElement(Main + "cellStyle", new XAttribute("name", name), new XAttribute("xfId", index),
            BuiltIn.TryGetValue(name, out var builtIn) ? new XAttribute("builtinId", builtIn) : null,
            BuiltIn.ContainsKey(name) ? new XAttribute("customBuiltin", 1) : null));
        _named[name] = index;
        _changed = true;
        return index;
    }

    private static void Set(XElement xf, string attribute, int value, string apply)
    {
        xf.SetAttributeValue(attribute, value);
        xf.SetAttributeValue(apply, 1);
    }

    private static void SetAlignment(XElement xf, CellStyle style)
    {
        var alignment = xf.Element(Main + "alignment");
        if (alignment is null)
        {
            alignment = new XElement(Main + "alignment");
            xf.AddFirst(alignment);
        }

        alignment.SetAttributeValue("horizontal", style.HorizontalAlignment == CellHorizontalAlignment.General ? null : Name(StylesReader.Horizontal, style.HorizontalAlignment));
        alignment.SetAttributeValue("vertical", style.VerticalAlignment == CellVerticalAlignment.Bottom ? null : Name(StylesReader.Vertical, style.VerticalAlignment));
        alignment.SetAttributeValue("wrapText", style.WrapText ? "1" : null);
        alignment.SetAttributeValue("textRotation", style.VerticalText ? 255 : style.TextRotation > 0 ? style.TextRotation : style.TextRotation < 0 ? 90 - style.TextRotation : null);
        alignment.SetAttributeValue("indent", style.Indent == 0 ? null : style.Indent.ToString(CultureInfo.InvariantCulture));
        alignment.SetAttributeValue("shrinkToFit", style.ShrinkToFit ? "1" : null);
        if (!alignment.HasAttributes && !alignment.HasElements)
            alignment.Remove();
        xf.SetAttributeValue("applyAlignment", 1);
    }

    private static void SetProtection(XElement xf, CellStyle style)
    {
        var protection = xf.Element(Main + "protection");
        if (protection is null)
        {
            protection = new XElement(Main + "protection");
            // Schema order: alignment, protection, extLst.
            if (xf.Element(Main + "alignment") is { } alignment)
                alignment.AddAfterSelf(protection);
            else
                xf.AddFirst(protection);
        }

        protection.SetAttributeValue("locked", style.Locked ? null : "0");
        protection.SetAttributeValue("hidden", style.FormulaHidden ? "1" : null);
        if (!protection.HasAttributes && !protection.HasElements)
            protection.Remove();
        xf.SetAttributeValue("applyProtection", 1);
    }

    private int Font(CellFont font)
    {
        var index = _read.Fonts.IndexOf(font);
        if (index >= 0)
            return index;
        return Add("fonts", FontElement(font, "font", "name"), _read.Fonts, font);
    }

    /// <summary>A font element; rich text runs call it rPr and name the font rFont.</summary>
    internal static XElement FontElement(CellFont font, string elementName, string nameElement)
    {
        var element = new XElement(Main + elementName);
        if (font.Bold)
            element.Add(new XElement(Main + "b"));
        if (font.Italic)
            element.Add(new XElement(Main + "i"));
        if (font.Strikethrough)
            element.Add(new XElement(Main + "strike"));
        if (font.Underline != CellUnderline.None)
            element.Add(new XElement(Main + "u", font.Underline == CellUnderline.Single ? null : new XAttribute("val", Name(StylesReader.Underlines, font.Underline))));
        element.Add(new XElement(Main + "sz", new XAttribute("val", font.Size.ToString("R", CultureInfo.InvariantCulture))));
        if (font.Script != CellFontScript.Normal)
            element.Add(new XElement(Main + "vertAlign", new XAttribute("val", font.Script == CellFontScript.Superscript ? "superscript" : "subscript")));
        if (font.Color is { } color)
            element.Add(Color("color", color));
        element.Add(new XElement(Main + nameElement, new XAttribute("val", font.Name)));
        return element;
    }

    // A fill equal to one of the file's is reused, apart from the reserved gray125 at index 1.
    private int Fill(FillLook fill)
    {
        for (var i = 0; i < _read.Fills.Count; i++)
        {
            if (i != 1 && _read.Fills[i] == fill)
                return i;
        }

        XElement content;
        if (fill.Gradient is { } gradient)
        {
            content = new XElement(Main + "gradientFill",
                gradient.Type == CellGradientType.Path ? new XAttribute("type", "path") : null,
                gradient.Type == CellGradientType.Linear && gradient.Angle != 0 ? new XAttribute("degree", Number(gradient.Angle)) : null,
                gradient.Type == CellGradientType.Path ? Edges() : null,
                gradient.Stops.Select(stop => new XElement(Main + "stop", new XAttribute("position", Number(stop.Position)), Color("color", stop.Color))));

            IEnumerable<XAttribute> Edges() =>
                new[] { ("left", gradient.Left), ("right", gradient.Right), ("top", gradient.Top), ("bottom", gradient.Bottom) }
                    .Where(edge => edge.Item2 != 0).Select(edge => new XAttribute(edge.Item1, Number(edge.Item2)));
        }
        else if (fill.Pattern == CellFillPattern.Solid)
        {
            content = fill.Color is { } color
                ? new XElement(Main + "patternFill", new XAttribute("patternType", "solid"), Color("fgColor", color), new XElement(Main + "bgColor", new XAttribute("indexed", 64)))
                : new XElement(Main + "patternFill", new XAttribute("patternType", "none"));
        }
        else
        {
            content = new XElement(Main + "patternFill", new XAttribute("patternType", Name(StylesReader.Patterns, fill.Pattern)),
                fill.PatternColor is { } pattern ? Color("fgColor", pattern) : null,
                fill.Color is { } background ? Color("bgColor", background) : null);
        }

        return Add("fills", new XElement(Main + "fill", content), _read.Fills, fill);

        static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
    }

    private int Border(BorderEdges edges)
    {
        var index = _read.Borders.IndexOf(edges);
        if (index >= 0)
            return index;
        return Add("borders", new XElement(Main + "border", Edge("left", edges.Left), Edge("right", edges.Right),
            Edge("top", edges.Top), Edge("bottom", edges.Bottom), new XElement(Main + "diagonal")), _read.Borders, edges);
    }

    private static XElement Edge(string name, CellBorder? border) => border is null
        ? new XElement(Main + name)
        : new XElement(Main + name, new XAttribute("style", Name(StylesReader.BorderStyles, border.Style)),
            border.Color is { } color ? Color("color", color) : null);

    private int NumberFormatId(string code)
    {
        if (code.Equals("General", StringComparison.OrdinalIgnoreCase))
            return 0;

        // 14 and 22 follow the reader's locale in Excel; a code typed as m/d/yyyy is kept as written.
        foreach (var (id, builtIn) in StylesReader.BuiltIn)
        {
            if (builtIn == code && id is not (14 or 22))
                return id;
        }

        foreach (var (id, custom) in _read.Custom)
        {
            if (custom == code)
                return id;
        }

        var next = Math.Max(163, _read.Custom.Count == 0 ? 0 : _read.Custom.Keys.Max()) + 1;
        _read.Custom[next] = code;
        Section("numFmts").Add(new XElement(Main + "numFmt", new XAttribute("numFmtId", next), new XAttribute("formatCode", code)));
        _changed = true;
        return next;
    }

    private int Add<T>(string section, XElement element, List<T> read, T value)
    {
        Section(section).Add(element);
        read.Add(value);
        _changed = true;
        return read.Count - 1;
    }

    // A section of the style sheet, created in schema order when the part has none.
    private XElement Section(string name)
    {
        var root = _document!.Root!;
        if (root.Element(Main + name) is { } section)
            return section;
        section = new XElement(Main + name);
        var rank = Array.IndexOf(Order, name);
        var previous = root.Elements().LastOrDefault(e => e.Name.Namespace == Main && Array.IndexOf(Order, e.Name.LocalName) is var r && r >= 0 && r < rank);
        if (previous is null)
            root.AddFirst(section);
        else
            previous.AddAfterSelf(section);
        _changed = true;
        return section;
    }

    /// <summary>The style as this file can hold it: theme colours as their RGB when the file has no theme part.</summary>
    public CellStyle Portable(CellStyle style) => _hasTheme ? style : style with
    {
        Font = Portable(style.Font),
        Fill = Plain(style.Fill),
        PatternColor = Plain(style.PatternColor),
        Gradient = style.Gradient is { } gradient
            ? new CellGradient(gradient.Stops.Select(stop => new GradientStop(stop.Position, Plain(stop.Color)!.Value)))
                { Type = gradient.Type, Angle = gradient.Angle, Left = gradient.Left, Right = gradient.Right, Top = gradient.Top, Bottom = gradient.Bottom }
            : null,
        LeftBorder = Plain(style.LeftBorder),
        RightBorder = Plain(style.RightBorder),
        TopBorder = Plain(style.TopBorder),
        BottomBorder = Plain(style.BottomBorder),
    };

    /// <summary>The font as this file can hold it.</summary>
    public CellFont Portable(CellFont font) => _hasTheme ? font : font with { Color = Plain(font.Color) };

    private static CellColor? Plain(CellColor? color) => color is { Theme: not null } value ? new CellColor(value.R, value.G, value.B) : color;

    private static CellBorder? Plain(CellBorder? border) => border is { Color: { Theme: not null } } ? border with { Color = Plain(border.Color) } : border;

    // A theme colour is saved as one, with its tint, so that it follows the theme in Excel.
    internal static XElement Color(string name, CellColor color) => color.Theme is { } theme
        ? new(Main + name, new XAttribute("theme", (int)theme), color.Tint == 0 ? null : new XAttribute("tint", color.Tint.ToString("R", CultureInfo.InvariantCulture)))
        : new(Main + name, new XAttribute("rgb", string.Create(CultureInfo.InvariantCulture, $"FF{color.R:X2}{color.G:X2}{color.B:X2}")));

    private static string Name<T>(Dictionary<string, T> map, T value) where T : struct =>
        map.First(pair => EqualityComparer<T>.Default.Equals(pair.Value, value)).Key;

    private static NumberFormat? ParseFormat(string code) =>
        !code.Equals("General", StringComparison.OrdinalIgnoreCase) && NumberFormat.TryParse(code, out var format) ? format : null;
}
