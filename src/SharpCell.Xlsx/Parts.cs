using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

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
    /// The inverse of <see cref="Decode"/>: carriage returns and characters XML cannot hold become
    /// <c>_xHHHH_</c>, and text that reads as such an escape gets its underscore escaped
    /// (<c>_x0041_</c> becomes <c>_x005F_x0041_</c>). Tabs and line feeds stay as they are.
    /// </summary>
    public static string Encode(string text)
    {
        StringBuilder? sb = null;
        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            var escape = ch == '_' && IsEscape(text, i) ? "_x005F_"
                : NeedsEscape(text, i) ? "_x" + ((int)ch).ToString("X4", CultureInfo.InvariantCulture) + "_"
                : null;
            if (escape is null)
            {
                sb?.Append(ch);
                continue;
            }

            sb ??= new StringBuilder(text.Length + 16).Append(text, 0, i);
            sb.Append(escape);
        }

        return sb?.ToString() ?? text;
    }

    // "_xHHHH_" at i, which Decode would read as one character.
    private static bool IsEscape(string text, int i) =>
        i + 7 <= text.Length && text[i + 1] == 'x' && text[i + 6] == '_'
        && ushort.TryParse(text.AsSpan(i + 2, 4), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out _);

    private static bool NeedsEscape(string text, int i)
    {
        var ch = text[i];
        if (ch < 0x20)
            return ch is not ('\t' or '\n');
        if (ch is '\uFFFE' or '\uFFFF')
            return true;
        if (char.IsHighSurrogate(ch))
            return i + 1 >= text.Length || !char.IsLowSurrogate(text[i + 1]);
        if (char.IsLowSurrogate(ch))
            return i == 0 || !char.IsHighSurrogate(text[i - 1]);
        return false;
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

/// <summary>What cells of a file point into, and reading a cell element the way the reader does.</summary>
internal sealed class SavedCells(IReadOnlyList<string> sharedStrings, CellMetadata? metadata)
{
    public IReadOnlyList<string> SharedStrings { get; } = sharedStrings;

    public CellMetadata? Metadata { get; } = metadata;

    /// <summary>The value saved in a &lt;c&gt; element, and whether it has a formula (&lt;f&gt;).</summary>
    public (CellValue Value, bool HasFormula) Read(XElement cell, Worksheet sheet, int row, int column)
    {
        using var reader = cell.CreateReader();
        reader.MoveToContent();
        var xml = WorksheetReader.ReadCell(reader);
        return (WorksheetReader.SavedValue(xml, this, sheet, new CellAddress(row, column)), xml.FormulaText is not null);
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
/// Cell metadata (<c>xl/metadata.xml</c>). A cell's <c>cm</c> attribute points at an entry that
/// says whether its array formula is a dynamic array (spills) rather than a Ctrl+Shift+Enter one;
/// its <c>vm</c> attribute points at a rich value, which is how newer errors such as <c>#SPILL!</c>
/// and <c>#CALC!</c> are saved (the cell itself says <c>#VALUE!</c>).
/// </summary>
internal sealed class CellMetadata
{
    private const string DynamicArrayType = "XLDAPR";
    private const string RichValueType = "XLRICHVALUE";

    private readonly List<bool> _dynamic = [];
    private readonly List<int> _richValues = [];
    private RichValues? _rich;

    /// <summary>Without a metadata part, any <c>cm</c> on an array formula is taken as dynamic.</summary>
    public static CellMetadata? Read(Package package, string? part, RichValues? rich)
    {
        if (part is null || !package.Exists(part))
            return null;

        var types = new List<string>();
        var futureDynamic = new List<bool>();
        var futureRich = new List<int>();
        var cells = new List<(int Type, int Value)>();
        var values = new List<(int Type, int Value)>();
        using var reader = package.OpenXml(part);
        string? section = null;
        string? futureName = null;
        var inBlock = false;
        var blockHasEntry = false;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement)
            {
                if (reader.LocalName is "futureMetadata" or "cellMetadata" or "valueMetadata" or "metadataTypes")
                    section = null;
                if (reader.LocalName == "bk")
                {
                    // Every block is one entry, even one without a record, so indexes stay aligned.
                    if (inBlock && !blockHasEntry)
                        BlockList(section, cells, values)?.Add((0, 0));
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
                    if (section == "futureMetadata" && futureName == RichValueType)
                        futureRich.Add(-1);
                    if (reader.IsEmptyElement)
                        BlockList(section, cells, values)?.Add((0, 0));
                    break;
                case "dynamicArrayProperties" when section == "futureMetadata" && futureDynamic.Count > 0:
                    futureDynamic[^1] = IsTrue(reader.GetAttribute("fDynamic"));
                    break;
                case "rvb" when section == "futureMetadata" && futureName == RichValueType && futureRich.Count > 0:
                    futureRich[^1] = ParseInt(reader.GetAttribute("i"));
                    break;
                case "rc" when inBlock && !blockHasEntry && BlockList(section, cells, values) is { } list:
                    blockHasEntry = true;
                    list.Add((ParseInt(reader.GetAttribute("t")), ParseInt(reader.GetAttribute("v"))));
                    break;
            }
        }

        bool Is(int type, string name) => type >= 1 && type <= types.Count && types[type - 1] == name;

        var metadata = new CellMetadata { _rich = rich };
        foreach (var (type, value) in cells)
            metadata._dynamic.Add(Is(type, DynamicArrayType) && value >= 0 && value < futureDynamic.Count && futureDynamic[value]);
        foreach (var (type, value) in values)
            metadata._richValues.Add(Is(type, RichValueType) && value >= 0 && value < futureRich.Count ? futureRich[value] : -1);
        return metadata;
    }

    private static List<(int, int)>? BlockList(string? section, List<(int, int)> cells, List<(int, int)> values) => section switch
    {
        "cellMetadata" => cells,
        "valueMetadata" => values,
        _ => null,
    };

    /// <param name="cm">The 1-based <c>cm</c> attribute of a cell.</param>
    public bool IsDynamicArray(int cm) => cm >= 1 && cm <= _dynamic.Count && _dynamic[cm - 1];

    /// <summary>The error a cell's rich value stands for, from its 1-based <c>vm</c> attribute; null if it is not an error.</summary>
    public ErrorKind? RichError(int vm) =>
        vm >= 1 && vm <= _richValues.Count && _rich is not null ? _rich.Error(_richValues[vm - 1]) : null;

    public static bool IsTrue(string? value) => value is "1" or "true";

    internal static int ParseInt(string? value) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : -1;
}

/// <summary>
/// Rich values (<c>xl/richData</c>). Only one kind matters for calculation: <c>_error</c>, whose
/// <c>errorType</c> is the ERROR.TYPE code minus one (13 is <c>#CALC!</c>, 8 is <c>#SPILL!</c>).
/// </summary>
internal sealed class RichValues
{
    private readonly List<ErrorKind?> _errors = [];

    public static RichValues? Read(Package package, string? valuesPart, string? structuresPart)
    {
        if (valuesPart is null || structuresPart is null || !package.Exists(valuesPart) || !package.Exists(structuresPart))
            return null;

        // For each structure: the position of its errorType key, or -1 if it is not an error.
        var errorKeyAt = new List<int>();
        using (var reader = package.OpenXml(structuresPart))
        {
            var key = 0;
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element)
                    continue;
                if (reader.LocalName == "s")
                {
                    errorKeyAt.Add(reader.GetAttribute("t") == "_error" ? -2 : -1);
                    key = 0;
                }
                else if (reader.LocalName == "k" && errorKeyAt.Count > 0)
                {
                    if (errorKeyAt[^1] == -2 && reader.GetAttribute("n") == "errorType")
                        errorKeyAt[^1] = key;
                    key++;
                }
            }
        }

        var rich = new RichValues();
        using (var reader = package.OpenXml(valuesPart))
        {
            reader.Read();
            while (!reader.EOF)
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "rv")
                {
                    reader.Read();
                    continue;
                }

                var structure = CellMetadata.ParseInt(reader.GetAttribute("s"));
                var errorKey = structure >= 0 && structure < errorKeyAt.Count ? errorKeyAt[structure] : -1;
                ErrorKind? error = null;
                var position = 0;
                using (var values = reader.ReadSubtree())
                {
                    // ReadElementContentAsString moves past the element, onto the next one: no Read after it.
                    values.Read();
                    while (!values.EOF)
                    {
                        if (values.NodeType != XmlNodeType.Element || values.LocalName != "v")
                        {
                            values.Read();
                            continue;
                        }

                        var text = values.ReadElementContentAsString();
                        if (position++ == errorKey && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var type)
                            && Enum.IsDefined((ErrorKind)(type + 1)))
                            error = (ErrorKind)(type + 1);
                    }
                }

                rich._errors.Add(error);
                reader.Read();
            }
        }

        return rich;
    }

    public ErrorKind? Error(int index) => index >= 0 && index < _errors.Count ? _errors[index] : null;
}
