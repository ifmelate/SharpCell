using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace SharpCell.Xlsx;

/// <summary>
/// The colours of <c>xl/theme/theme1.xml</c> in the order cells refer to them: Excel's theme index
/// 0 is lt1 and 1 is dk1, although the part lists dk1 first; 4 to 9 are the accents, 10 and 11 the
/// hyperlink colours.
/// </summary>
internal static class ThemeReader
{
    private const string DrawingMl = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly string[] Order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

    /// <summary>The Office theme of Excel 2013 to 2022, for files without a theme part.</summary>
    public static CellColor[] Office => Workbook.OfficeTheme;

    /// <summary>The theme's colours; the Office theme's for any the part lacks or cannot give.</summary>
    public static CellColor[] Read(Package package, string? part)
    {
        var colors = (CellColor[])Office.Clone();
        if (part is null || !package.Exists(part))
            return colors;
        try
        {
            using var reader = package.OpenXml(part);
            var scheme = XDocument.Load(reader).Descendants(XName.Get("clrScheme", DrawingMl)).FirstOrDefault();
            if (scheme is null)
                return colors;
            for (var i = 0; i < Order.Length; i++)
            {
                var entry = scheme.Element(XName.Get(Order[i], DrawingMl));
                var srgb = entry?.Element(XName.Get("srgbClr", DrawingMl))?.Attribute("val")?.Value;
                var system = entry?.Element(XName.Get("sysClr", DrawingMl))?.Attribute("lastClr")?.Value;
                if (StylesReader.ParseRgb(srgb ?? system) is { } color)
                    colors[i] = color;
            }
        }
        catch (XmlException)
        {
        }

        return colors;
    }
}
