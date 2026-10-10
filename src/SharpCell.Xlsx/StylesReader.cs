using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using SharpCell.Functions;

namespace SharpCell.Xlsx;

/// <summary>
/// The number format of each cell style in <c>xl/styles.xml</c>: a cell's <c>s</c> attribute is an
/// index into <c>cellXfs</c>, whose <c>numFmtId</c> is a custom code from <c>numFmts</c> or one of
/// Excel's built-in formats.
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

    /// <summary>
    /// The format of each <c>cellXfs</c> entry, in order; null for General and for a code SharpCell
    /// cannot read, so a file is never refused over a format.
    /// </summary>
    public static List<NumberFormat?> Read(Package package, string part)
    {
        var custom = new Dictionary<int, string>();
        var styles = new List<NumberFormat?>();
        using var reader = package.OpenXml(part);
        var inCellXfs = false;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "cellXfs")
                inCellXfs = false;
            if (reader.NodeType != XmlNodeType.Element)
                continue;

            switch (reader.LocalName)
            {
                case "numFmt" when Id(reader) is { } id && reader.GetAttribute("formatCode") is { } code:
                    custom[id] = code;
                    break;
                case "cellXfs":
                    inCellXfs = !reader.IsEmptyElement;
                    break;
                case "xf" when inCellXfs:
                    styles.Add(Format(Id(reader) ?? 0, custom));
                    break;
            }
        }

        return styles;
    }

    private static int? Id(XmlReader reader) =>
        int.TryParse(reader.GetAttribute("numFmtId"), NumberStyles.None, CultureInfo.InvariantCulture, out var id) ? id : null;

    private static NumberFormat? Format(int id, Dictionary<int, string> custom)
    {
        if (!custom.TryGetValue(id, out var code) && !BuiltIn.TryGetValue(id, out code))
            return null;
        if (code.Equals("General", System.StringComparison.OrdinalIgnoreCase))
            return null;
        return NumberFormat.TryParse(code, out var format) ? format : null;
    }
}
