using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

/// <summary>Conditional formats are read from files with their differential formats and saved back.</summary>
public class XlsxConditionalFormatTests
{
    private const string StylesXml = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
        <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
        <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
        <dxfs count="2">
        <dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
        <dxf><font><b/><i val="0"/></font><numFmt numFmtId="164" formatCode="0.0%"/><border><bottom style="thin"><color rgb="FF000000"/></bottom></border></dxf>
        </dxfs>
        </styleSheet>
        """;

    // A1:A5 hold 1..5; rules: a colour scale (priority 3), a cell value rule (1), a text rule (2) and
    // an above average rule over two ranges (4).
    private static TestXlsx Book() => new TestXlsx
    {
        Styles = StylesXml,
    }.RawSheet("S", """
        <sheetData>
        <row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>apple pie</t></is></c></row>
        <row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>
        <row r="4"><c r="A4"><v>4</v></c></row><row r="5"><c r="A5"><v>5</v></c></row>
        </sheetData>
        <conditionalFormatting sqref="A1:A5"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>
        <cfRule type="cellIs" dxfId="0" priority="1" stopIfTrue="1" operator="greaterThan"><formula>$C$1+3</formula></cfRule></conditionalFormatting>
        <conditionalFormatting sqref="B1:B5"><cfRule type="containsText" dxfId="1" priority="2" operator="containsText" text="pie"><formula>NOT(ISERROR(SEARCH("pie",B1)))</formula></cfRule></conditionalFormatting>
        <conditionalFormatting sqref="A1:A2 A4"><cfRule type="aboveAverage" dxfId="0" priority="4" aboveAverage="0" equalAverage="1"/></conditionalFormatting>
        <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        """);

    private static (byte[] Original, Workbook Workbook) Load()
    {
        var original = Book().Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        return (original, workbook);
    }

    private static (string Sheet, string Styles, Workbook Reloaded) SaveChecked(byte[] original, Workbook workbook)
    {
        var saved = Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var reloaded = Reload(saved);
        Assert.Equal(Describe(workbook["S"]), Describe(reloaded["S"]));
        return (Part(saved, "xl/worksheets/sheet1.xml"), Part(saved, "xl/styles.xml"), reloaded);
    }

    private static List<string> Describe(Worksheet sheet) =>
        [.. sheet.ConditionalFormats.Select(f => $"{string.Join(" ", f.Ranges.Select(r => r.Address))} {f.Format} {f.Format.Style} {f.Format.StopIfTrue}")];

    [Fact]
    public void Rules_are_read_in_priority_order_with_their_looks()
    {
        var (_, workbook) = Load();
        var sheet = workbook["S"];

        Assert.Equal(
            [ConditionalFormatType.CellValue, ConditionalFormatType.ContainsText, ConditionalFormatType.ColorScale, ConditionalFormatType.BelowAverage],
            sheet.ConditionalFormats.Select(f => f.Format.Type));
        var greater = sheet.ConditionalFormats[0].Format;
        Assert.Equal((ConditionalOperator.GreaterThan, "=$C$1+3", true), (greater.Operator, greater.Formula1, greater.StopIfTrue));
        Assert.Equal(ConditionalStyle.LightRedFill, greater.Style);
        Assert.Equal(new ConditionalStyle { Bold = true, Italic = false, NumberFormat = "0.0%", BottomBorder = new CellBorder(CellBorderStyle.Thin, CellColor.FromRgb(0)) },
            sheet.ConditionalFormats[1].Format.Style);
        Assert.Equal("pie", sheet.ConditionalFormats[1].Format.Text);
        Assert.Null(sheet.ConditionalFormats[2].Format.Style);
        Assert.Equal(["A1:A2", "A4"], sheet.ConditionalFormats[3].Ranges.Select(r => r.Address));
        Assert.True(sheet.ConditionalFormats[3].Format.OrEqual);
    }

    [Fact]
    public void Rules_read_from_a_file_tell_which_cells_they_apply_to()
    {
        var (_, workbook) = Load();
        var sheet = workbook["S"];

        Assert.Equal([ConditionalFormatType.CellValue], sheet["A4"].ConditionalFormats.Select(f => f.Type));   // stop if true
        Assert.Equal([ConditionalFormatType.ColorScale, ConditionalFormatType.BelowAverage], sheet["A1"].ConditionalFormats.Select(f => f.Type));
        Assert.True(sheet["B1"].DisplayStyle.Font.Bold);
        Assert.Equal(CellColor.FromRgb(0xFFC7CE), sheet["A5"].DisplayStyle.Fill);
    }

    [Fact]
    public void Unchanged_rules_stay_as_the_file_has_them()
    {
        var (original, workbook) = Load();
        workbook["S"]["A1"].Value = 7;

        var (sheet, _, _) = SaveChecked(original, workbook);

        Assert.Contains("<cfRule type=\"colorScale\" priority=\"3\">", sheet);
        Assert.Contains("priority=\"4\" aboveAverage=\"0\"", sheet);
    }

    [Fact]
    public void A_new_rule_is_saved_with_a_new_look_and_the_file_rules_keep_theirs()
    {
        var (original, workbook) = Load();
        var sheet = workbook["S"];
        sheet.AddConditionalFormat("C1:C9", ConditionalFormat.Top(3, new ConditionalStyle { Fill = CellColor.FromRgb(0x00B050) }, percent: true));
        sheet.AddConditionalFormat("D1:D9", ConditionalFormat.CellValue(ConditionalOperator.Between, "=1", ConditionalStyle.LightRedFill, "=$E$1"));

        var (part, styles, _) = SaveChecked(original, workbook);

        Assert.Contains("<color rgb=\"FF63BE7B\"", part);                                      // the colour scale kept its settings
        Assert.Contains("type=\"colorScale\" priority=\"3\"", part);
        Assert.Contains("type=\"top10\" dxfId=\"2\" priority=\"5\" rank=\"3\" percent=\"1\"", part);
        Assert.Contains("type=\"cellIs\" dxfId=\"0\" priority=\"6\" operator=\"between\"><formula>1</formula><formula>$E$1</formula>", part);
        Assert.Contains("<dxfs count=\"3\">", styles);
        Assert.Contains("<bgColor rgb=\"FF00B050\"", styles);
    }

    [Fact]
    public void Rules_follow_inserted_rows_and_cleared_ones_go()
    {
        var (original, workbook) = Load();
        var sheet = workbook["S"];
        sheet.InsertRows(1);
        sheet.ClearConditionalFormats("B1:B10");

        var (part, _, reloaded) = SaveChecked(original, workbook);

        Assert.Contains("<conditionalFormatting sqref=\"A2:A6\"", part);
        Assert.Contains("<formula>$C$2+3</formula>", part);
        Assert.DoesNotContain("containsText", part);
        Assert.Equal(["A2:A3", "A5"], reloaded["S"].ConditionalFormats[^1].Ranges.Select(r => r.Address));
    }

    [Fact]
    public void A_workbook_built_in_code_saves_its_rules()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = 1;
        sheet.AddConditionalFormat("A1:A10", ConditionalFormat.Expression("=A1>AVERAGE($A$1:$A$10)", new ConditionalStyle { Bold = true, FontColor = CellColor.FromRgb(0xFF0000) }));
        sheet.AddConditionalFormat("B1:B10", ConditionalFormat.DateOccurring(TimePeriod.ThisMonth, ConditionalStyle.YellowFill) with { StopIfTrue = true });
        sheet.AddConditionalFormat("C1:C10", ConditionalFormat.NotContainsText("x\"y", ConditionalStyle.GreenFill));
        sheet.AddConditionalFormat("D1:D10", ConditionalFormat.DuplicateValues(new ConditionalStyle { NumberFormat = "0.00", Underline = CellUnderline.Double }));

        var saved = Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        Assert.Equal(Describe(sheet), Describe(Reload(saved)["S"]));
        Assert.Contains("<formula>AND(MONTH(B1)=MONTH(TODAY()),YEAR(B1)=YEAR(TODAY()))</formula>", Part(saved, "xl/worksheets/sheet1.xml"));
    }
}
