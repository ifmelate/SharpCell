using System.Globalization;
using System.Xml.Linq;

namespace SharpCell.Xlsx;

/// <summary>How values are written into a cell element (&lt;c&gt;).</summary>
internal static class CellEncoding
{
    /// <summary>Values a cell cannot hold (arrays, functions) become #CALC!, as a formula result does.</summary>
    public static CellValue Storable(CellValue value) =>
        value.Kind is CellValueKind.Array or CellValueKind.Lambda or CellValueKind.Missing ? CellValue.Error(ErrorKind.Calc) : value;

    /// <summary>The cell's t attribute; null for a number or an empty value.</summary>
    public static string? TypeOf(CellValue value, bool isFormula) => value.Kind switch
    {
        CellValueKind.Text => isFormula ? "str" : "inlineStr",
        CellValueKind.Boolean => "b",
        CellValueKind.Error => "e",
        _ => null,
    };

    /// <summary>
    /// The element holding the value: &lt;v&gt;, or &lt;is&gt; for a constant's text (the shared
    /// string table is left alone); null for an empty value. #SPILL! and #CALC! are written as
    /// #VALUE!, as files show them to older programs; the cell's vm attribute names the real error.
    /// </summary>
    public static XElement? ValueElement(XNamespace ns, CellValue value, bool isFormula) => value.Kind switch
    {
        CellValueKind.Number => new XElement(ns + "v", FormatNumber(value.AsNumber())),
        CellValueKind.Text when isFormula => new XElement(ns + "v", XmlText.Encode(value.AsText())),
        CellValueKind.Text => new XElement(ns + "is",
            new XElement(ns + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), XmlText.Encode(value.AsText()))),
        CellValueKind.Boolean => new XElement(ns + "v", value.AsBoolean() ? "1" : "0"),
        CellValueKind.Error => new XElement(ns + "v", IsRichError(value.AsError()) ? "#VALUE!" : value.AsError().ToText()),
        _ => null,
    };

    /// <summary>Errors files store as rich values behind #VALUE!.</summary>
    public static bool IsRichError(ErrorKind kind) => kind is ErrorKind.Spill or ErrorKind.Calc;

    /// <summary>The shortest text that reads back as the same double; negative zero is 0.</summary>
    public static string FormatNumber(double value) => value == 0 ? "0" : value.ToString("R", CultureInfo.InvariantCulture);
}
