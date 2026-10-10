using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class StylesReaderTests
{
    private const string Styles = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <numFmts count="1"><numFmt numFmtId="164" formatCode="0.0&quot;x&quot;"/></numFmts>
          <cellXfs count="5">
            <xf numFmtId="0"/><xf numFmtId="4" applyNumberFormat="1"/><xf numFmtId="164"/><xf numFmtId="14"/><xf numFmtId="100"/>
          </cellXfs>
        </styleSheet>
        """;

    private static Workbook Load(string sheet, string? styles = Styles)
    {
        var file = new TestXlsx { Styles = styles }.Sheet("S", sheet);
        return XlsxReader.Load(file.Build());
    }

    [Fact]
    public void Built_in_and_custom_formats_come_from_the_cell_style()
    {
        var sheet = Load("<row r=\"1\"><c r=\"A1\" s=\"1\"><v>1</v></c><c r=\"B1\" s=\"2\"><v>1</v></c><c r=\"C1\" s=\"3\"><v>1</v></c><c r=\"D1\"><v>1</v></c><c r=\"E1\" s=\"0\"><v>1</v></c><c r=\"F1\" s=\"4\"><v>1</v></c></row>")["S"];
        Assert.Equal("#,##0.00", sheet["A1"].NumberFormat);
        Assert.Equal("0.0\"x\"", sheet["B1"].NumberFormat);
        Assert.Equal("m/d/yyyy", sheet["C1"].NumberFormat);
        Assert.Equal("General", sheet["D1"].NumberFormat);
        Assert.Equal("General", sheet["E1"].NumberFormat);
        Assert.Equal("General", sheet["F1"].NumberFormat);   // an id below 164 with no built-in code
    }

    [Fact]
    public void A_formula_cell_and_an_empty_styled_cell_keep_their_format()
    {
        var sheet = Load("<row r=\"1\"><c r=\"A1\" s=\"1\"><f>1+1</f><v>2</v></c><c r=\"B1\" s=\"1\"/></row>")["S"];
        Assert.Equal("#,##0.00", sheet["A1"].NumberFormat);
        Assert.Equal("#,##0.00", sheet["B1"].NumberFormat);
        Assert.Equal(CellValue.Empty, sheet["B1"].Value);
    }

    [Fact]
    public void No_styles_part_means_General_everywhere()
    {
        Assert.Equal("General", Load("<row r=\"1\"><c r=\"A1\" s=\"5\"><v>1</v></c></row>", styles: null)["S"]["A1"].NumberFormat);
    }

    [Fact]
    public void A_style_index_outside_cellXfs_is_invalid_data()
    {
        Assert.Throws<System.IO.InvalidDataException>(() => Load("<row r=\"1\"><c r=\"A1\" s=\"9\"><v>1</v></c></row>"));
    }

    [Fact]
    public void An_unparsable_custom_code_falls_back_to_General()
    {
        var styles = Styles.Replace("0.0&quot;x&quot;", "0;0;0;0;0");
        Assert.Equal("General", Load("<row r=\"1\"><c r=\"A1\" s=\"2\"><v>1</v></c></row>", styles)["S"]["A1"].NumberFormat);
    }

    [Fact]
    public void Every_corpus_file_still_loads()
    {
        foreach (var relative in CorpusFiles.All)
        {
            try
            {
                XlsxReader.Load(CorpusFiles.PathOf(relative));
            }
            catch (NotSupportedException)
            {
            }
        }
    }
}
