using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;

namespace SharpCell.Xlsx;

internal static class XmlText
{
    /// <summary>
    /// Undoes the escaping of characters XML cannot hold: <c>_x000D_</c> is a carriage return and
    /// <c>_x005F_</c> a literal underscore (so <c>_x005F_x0041_</c> reads as <c>_x0041_</c>).
    /// </summary>
    public static string Decode(string text)
    {
        var index = text.IndexOf("_x", StringComparison.Ordinal);
        if (index < 0)
            return text;

        var sb = new StringBuilder(text.Length);
        var start = 0;
        while (index >= 0)
        {
            if (index + 7 <= text.Length && text[index + 6] == '_'
                && ushort.TryParse(text.AsSpan(index + 2, 4), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var code))
            {
                sb.Append(text, start, index - start).Append((char)code);
                start = index + 7;
                index = text.IndexOf("_x", start, StringComparison.Ordinal);
            }
            else
            {
                index = text.IndexOf("_x", index + 1, StringComparison.Ordinal);
            }
        }

        return sb.Append(text, start, text.Length - start).ToString();
    }

    /// <summary>
    /// Text of a string item (<c>si</c> or <c>is</c>) the reader is on: plain <c>t</c> or rich text
    /// runs concatenated; phonetic guides (<c>rPh</c>) are not part of the value. Leaves the reader
    /// after the item.
    /// </summary>
    public static string ReadStringItem(XmlReader reader)
    {
        if (reader.IsEmptyElement)
        {
            reader.Read();
            return "";
        }

        var depth = reader.Depth;
        var sb = new StringBuilder();
        reader.Read();
        while (!(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) && !reader.EOF)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "rPh")
            {
                reader.Skip();
                continue;
            }

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t")
            {
                sb.Append(reader.ReadElementContentAsString());
                continue;
            }

            reader.Read();
        }

        reader.Read();
        return Decode(sb.ToString());
    }
}

internal static class SharedStrings
{
    public static List<string> Read(Package package, string part)
    {
        var strings = new List<string>();
        using var reader = package.OpenXml(part);
        reader.Read();
        while (!reader.EOF)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
                strings.Add(XmlText.ReadStringItem(reader));
            else
                reader.Read();
        }

        return strings;
    }
}

/// <summary>
/// Cell metadata (<c>xl/metadata.xml</c>): a cell's <c>cm</c> attribute points at an entry that
/// says whether its array formula is a dynamic array (spills) rather than a Ctrl+Shift+Enter one.
/// </summary>
internal sealed class CellMetadata
{
    private const string DynamicArrayType = "XLDAPR";

    private readonly List<bool> _dynamic = [];

    /// <summary>Without a metadata part, any <c>cm</c> on an array formula is taken as dynamic.</summary>
    public static CellMetadata? Read(Package package, string? part)
    {
        if (part is null || !package.Exists(part))
            return null;

        var types = new List<string>();
        var futureDynamic = new List<bool>();
        var cells = new List<(int Type, int Value)>();
        using var reader = package.OpenXml(part);
        string? section = null;
        string? futureName = null;
        var inBlock = false;
        var blockHasEntry = false;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement)
            {
                if (reader.LocalName is "futureMetadata" or "cellMetadata" or "metadataTypes")
                    section = null;
                if (reader.LocalName == "bk")
                {
                    // Every block is one entry, even one without a record, so indexes stay aligned.
                    if (section == "cellMetadata" && inBlock && !blockHasEntry)
                        cells.Add((0, 0));
                    inBlock = false;
                }
                continue;
            }

            if (reader.NodeType != XmlNodeType.Element)
                continue;

            switch (reader.LocalName)
            {
                case "metadataTypes" or "cellMetadata" or "valueMetadata":
                    section = reader.LocalName;
                    break;
                case "futureMetadata":
                    section = reader.LocalName;
                    futureName = reader.GetAttribute("name");
                    break;
                case "metadataType" when section == "metadataTypes":
                    types.Add(reader.GetAttribute("name") ?? "");
                    break;
                case "bk":
                    inBlock = !reader.IsEmptyElement;
                    blockHasEntry = false;
                    if (section == "futureMetadata" && futureName == DynamicArrayType)
                        futureDynamic.Add(false);
                    if (section == "cellMetadata" && reader.IsEmptyElement)
                        cells.Add((0, 0));
                    break;
                case "dynamicArrayProperties" when section == "futureMetadata" && futureDynamic.Count > 0:
                    futureDynamic[^1] = IsTrue(reader.GetAttribute("fDynamic"));
                    break;
                case "rc" when section == "cellMetadata" && inBlock && !blockHasEntry:
                    blockHasEntry = true;
                    cells.Add((ParseInt(reader.GetAttribute("t")), ParseInt(reader.GetAttribute("v"))));
                    break;
            }
        }

        var metadata = new CellMetadata();
        foreach (var (type, value) in cells)
        {
            var isDynamic = type >= 1 && type <= types.Count && types[type - 1] == DynamicArrayType
                && value >= 0 && value < futureDynamic.Count && futureDynamic[value];
            metadata._dynamic.Add(isDynamic);
        }

        return metadata;
    }

    /// <param name="cm">The 1-based <c>cm</c> attribute of a cell.</param>
    public bool IsDynamicArray(int cm) => cm >= 1 && cm <= _dynamic.Count && _dynamic[cm - 1];

    public static bool IsTrue(string? value) => value is "1" or "true";

    private static int ParseInt(string? value) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : -1;
}
