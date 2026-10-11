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
    private string? _part;
    private XDocument? _document;
    private StyleSheet _read = new();
    private List<XElement> _xfs = [];
    private bool _changed;
    private readonly Dictionary<(int, CellStyle, string), int> _cache = [];

    private StylesWriter(PackageEdit edit, string workbookPart, string? part, CellColor[] theme)
    {
        _edit = edit;
        _workbookPart = workbookPart;
        _part = part;
        _theme = theme;
    }

    public static StylesWriter Open(PackageEdit edit, string workbookPart, Dictionary<string, Relationship> relationships)
    {
        var part = WorkbookReader.FirstOfType(relationships, "styles")?.Target;
        var theme = ThemeReader.Read(edit.Package, WorkbookReader.FirstOfType(relationships, "theme")?.Target);
        return new StylesWriter(edit, workbookPart, part is not null && edit.Package.Exists(part) ? part : null, theme);
    }

    /// <summary>The index of a cell style with this look and format, based on the style at <paramref name="baseIndex"/>.</summary>
    /// <param name="baseIndex">The cell's style in the file, or 0.</param>
    /// <param name="style">How the cell looks.</param>
    /// <param name="formatCode">The number format code; <c>General</c> for none.</param>
    public int XfFor(int baseIndex, CellStyle style, string formatCode)
    {
        if (_cache.TryGetValue((baseIndex, style, formatCode), out var cached))
            return cached;
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

        _cache[(baseIndex, style, formatCode)] = index;
        return index;
    }

    /// <summary>Changes the first cell style, which cells without a style of their own show, and the Normal style.</summary>
    public void SetDefault(CellStyle style)
    {
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
            if (section.Attribute("count") is not null || section.Name.LocalName is "numFmts" or "fonts" or "fills" or "borders" or "cellXfs")
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
        if (_xfs.Count == 0)
        {
            var xf = new XElement(Main + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", Font(CellFont.Default)),
                new XAttribute("fillId", Fill(null)), new XAttribute("borderId", Border(BorderEdges.None)));
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
        if (style.Fill != was.Fill)
            Set(xf, "fillId", Fill(style.Fill), "applyFill");
        var edges = new BorderEdges(style.LeftBorder, style.RightBorder, style.TopBorder, style.BottomBorder);
        if (edges != new BorderEdges(was.LeftBorder, was.RightBorder, was.TopBorder, was.BottomBorder))
            Set(xf, "borderId", Border(edges), "applyBorder");
        if ((style.HorizontalAlignment, style.VerticalAlignment, style.WrapText, style.Indent)
            != (was.HorizontalAlignment, was.VerticalAlignment, was.WrapText, was.Indent))
            SetAlignment(xf, style);
        if ((style.Locked, style.FormulaHidden) != (was.Locked, was.FormulaHidden))
            SetProtection(xf, style);
        if (formatCode != (old.Format?.FormatCode ?? "General"))
            Set(xf, "numFmtId", NumberFormatId(formatCode), "applyNumberFormat");
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
        alignment.SetAttributeValue("indent", style.Indent == 0 ? null : style.Indent.ToString(CultureInfo.InvariantCulture));
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
        var element = new XElement(Main + "font");
        if (font.Bold)
            element.Add(new XElement(Main + "b"));
        if (font.Italic)
            element.Add(new XElement(Main + "i"));
        if (font.Strikethrough)
            element.Add(new XElement(Main + "strike"));
        if (font.Underline != CellUnderline.None)
            element.Add(new XElement(Main + "u", font.Underline == CellUnderline.Single ? null : new XAttribute("val", Name(StylesReader.Underlines, font.Underline))));
        element.Add(new XElement(Main + "sz", new XAttribute("val", font.Size.ToString("R", CultureInfo.InvariantCulture))));
        if (font.Color is { } color)
            element.Add(Color("color", color));
        element.Add(new XElement(Main + "name", new XAttribute("val", font.Name)));
        return Add("fonts", element, _read.Fonts, font);
    }

    private int Fill(CellColor? color)
    {
        if (color is null)
        {
            var empty = _read.EmptyFills.IndexOf(true);
            if (empty >= 0)
                return empty;
            _read.EmptyFills.Add(true);
            return Add("fills", new XElement(Main + "fill", new XElement(Main + "patternFill", new XAttribute("patternType", "none"))), _read.Fills, null);
        }

        var index = _read.Fills.IndexOf(color);
        if (index >= 0)
            return index;
        _read.EmptyFills.Add(false);
        return Add("fills", new XElement(Main + "fill", new XElement(Main + "patternFill", new XAttribute("patternType", "solid"),
            Color("fgColor", color.Value), new XElement(Main + "bgColor", new XAttribute("indexed", 64)))), _read.Fills, color);
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

    private static XElement Color(string name, CellColor color) =>
        new(Main + name, new XAttribute("rgb", string.Create(CultureInfo.InvariantCulture, $"FF{color.R:X2}{color.G:X2}{color.B:X2}")));

    private static string Name<T>(Dictionary<string, T> map, T value) where T : struct =>
        map.First(pair => EqualityComparer<T>.Default.Equals(pair.Value, value)).Key;

    private static NumberFormat? ParseFormat(string code) =>
        !code.Equals("General", StringComparison.OrdinalIgnoreCase) && NumberFormat.TryParse(code, out var format) ? format : null;
}
