using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using SharpCell.Xlsx;
using XmlText = SharpCell.Xlsx.XmlText;

namespace SharpCell.Tests.Xlsx;

public class CellEncodingTests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Theory]
    [InlineData("plain")]
    [InlineData("a\rb")]
    [InlineData("tab\tand\nnewline")]
    [InlineData("\u0001\u001F")]
    [InlineData("_x0041_")]
    [InlineData("_x005F_")]
    [InlineData("_x00e9_")]
    [InlineData("x_y _x _x12_")]
    [InlineData("emoji 😀")]
    [InlineData("￿￾")]
    public void Encoded_text_decodes_to_itself_and_is_valid_XML(string text)
    {
        var encoded = XmlText.Encode(text);

        Assert.Equal(text, XmlText.Decode(encoded));
        XmlConvert.VerifyXmlChars(encoded);
    }

    [Theory]
    [InlineData("a\rb", "a_x000D_b")]
    [InlineData("_x0041_", "_x005F_x0041_")]
    [InlineData("\u0001", "_x0001_")]
    public void Escapes_what_XML_or_Decode_would_change(string text, string expected) =>
        Assert.Equal(expected, XmlText.Encode(text));

    // Built in code: a lone surrogate in an attribute argument does not survive compilation.
    [Fact]
    public void Lone_surrogates_are_escaped()
    {
        var high = new string((char)0xD800, 1) + " lone";
        var low = new string((char)0xDC00, 1) + " lone";

        Assert.Equal("_xD800_ lone", XmlText.Encode(high));
        Assert.Equal("_xDC00_ lone", XmlText.Encode(low));
        Assert.Equal(high, XmlText.Decode(XmlText.Encode(high)));
        Assert.Equal(low, XmlText.Decode(XmlText.Encode(low)));
    }

    [Fact]
    public void Plain_text_is_returned_as_it_is()
    {
        var text = "nothing to escape";
        Assert.Same(text, XmlText.Encode(text));
    }

    public static TheoryData<string, bool, string?, string> Values() => new()
    {
        { "number", false, null, "<v>1.5</v>" },
        { "negative zero", true, null, "<v>0</v>" },
        { "tiny", true, null, "<v>1E-300</v>" },
        { "text constant", false, "inlineStr", "<is><t xml:space=\"preserve\"> a b_x000D_ </t></is>" },
        { "text result", true, "str", "<v> a b_x000D_ </v>" },
        { "true", true, "b", "<v>1</v>" },
        { "false", false, "b", "<v>0</v>" },
        { "div0", true, "e", "<v>#DIV/0!</v>" },
        { "spill", true, "e", "<v>#VALUE!</v>" },
        { "calc", true, "e", "<v>#VALUE!</v>" },
    };

    [Theory]
    [MemberData(nameof(Values))]
    public void Writes_the_type_and_value_element(string name, bool isFormula, string? type, string xml)
    {
        var value = name switch
        {
            "number" => CellValue.Number(1.5),
            "negative zero" => CellValue.Number(-0.0),
            "tiny" => CellValue.Number(1e-300),
            "text constant" or "text result" => CellValue.Text(" a b\r "),
            "true" => CellValue.True,
            "false" => CellValue.False,
            "div0" => CellValue.Error(ErrorKind.Div0),
            "spill" => CellValue.Error(ErrorKind.Spill),
            _ => CellValue.Error(ErrorKind.Calc),
        };

        Assert.Equal(type, CellEncoding.TypeOf(value, isFormula));
        var element = CellEncoding.ValueElement(Ns, value, isFormula)!;
        Assert.Equal(xml, element.ToString(SaveOptions.DisableFormatting).Replace(" xmlns=\"" + Ns.NamespaceName + "\"", ""));
    }

    [Fact]
    public void Values_a_cell_cannot_hold_become_calc()
    {
        Assert.Equal(CellValue.Error(ErrorKind.Calc), CellEncoding.Storable(CellValue.Array(new CellValue[1, 1])));
        Assert.Equal(CellValue.Number(2), CellEncoding.Storable(CellValue.Number(2)));
        Assert.Null(CellEncoding.ValueElement(Ns, CellValue.Empty, isFormula: false));
    }

    [Fact]
    public void Numbers_read_back_exactly()
    {
        foreach (var number in new[] { 0.1 + 0.2, 1e15 + 0.3, -123456.789, double.Epsilon, 45301 })
            Assert.Equal(number, double.Parse(CellEncoding.FormatNumber(number), CultureInfo.InvariantCulture));
    }
}
