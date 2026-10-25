using System.IO.Compression;
using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

/// <summary>Sheets renamed, removed, moved, hidden and coloured are saved into the file they were read from.</summary>
public class XlsxSheetOperationsTests
{
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string ChartType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

    // Sheet B has a note, a table and a drawing with a chart; A's formulas, a chart and names point at B.
    private static TestXlsx Book() => new TestXlsx { CalcChain = ["A1"], WorkbookExtra = """<bookViews><workbookView activeTab="1"/></bookViews>""" }
        .RawSheet("A", """<sheetData><row r="1"><c r="A1"><f>B!A1*2</f><v>20</v></c><c r="B1"><f>SUM(A:C!A1)</f><v>111</v></c></row></sheetData>""")
        .RawSheet("B", """
            <sheetData><row r="1"><c r="A1"><v>10</v></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Units</t></is></c></row><row r="3"><c r="A3"><v>5</v></c></row></sheetData>
            <drawing r:id="rD"/><legacyDrawing r:id="rV"/><tableParts count="1"><tablePart r:id="rT"/></tableParts>
            """)
        .RawSheet("C", """<sheetData><row r="1"><c r="A1"><v>100</v></c></row></sheetData>""")
        .SheetRelationships("B", $"""
            <Relationship Id="rV" Type="{R}vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
            <Relationship Id="rC" Type="{R}comments" Target="../comments1.xml"/>
            <Relationship Id="rD" Type="{R}drawing" Target="../drawings/drawing1.xml"/>
            <Relationship Id="rT" Type="{R}table" Target="../tables/table1.xml"/>
            """)
        .Part("xl/comments1.xml", """
            <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Ann</author></authors><commentList>
            <comment ref="A1" authorId="0"><text><t>plain</t></text></comment></commentList></comments>
            """, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
        .Part("xl/drawings/vmlDrawing1.vml", """
            <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
            <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;visibility:hidden"><x:ClientData ObjectType="Note"><x:Anchor>1, 15, 0, 10, 3, 15, 4, 4</x:Anchor><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>
            </xml>
            """, "application/vnd.openxmlformats-officedocument.vmlDrawing")
        .Part("xl/tables/table1.xml", """
            <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A2:A3"><autoFilter ref="A2:A3"/>
            <tableColumns count="1"><tableColumn id="1" name="Units"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>
            """, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
        .Part("xl/drawings/drawing1.xml", """
            <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
            """, "application/vnd.openxmlformats-officedocument.drawing+xml")
        .Part("xl/drawings/_rels/drawing1.xml.rels", $"""
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rCh" Type="{R}chart" Target="../charts/chart1.xml"/></Relationships>
            """)
        .Part("xl/charts/chart1.xml", """
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>B!$A$1:$A$3</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>
            """, ChartType)
        .Name("Ten", "B!$A$1")
        .Name("Local", "C!$A$1", " localSheetId=\"2\"")
        .Name("_xlnm.Print_Area", "B!$A$1:$A$3", " localSheetId=\"1\"");

    private static (byte[] Original, Workbook Workbook) Load()
    {
        var original = Book().Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        return (original, workbook);
    }

    private static byte[] SaveChecked(byte[] original, Workbook workbook)
    {
        workbook.Recalculate();
        var saved = Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var reloaded = Reload(saved);
        Assert.Equal(Structure.Of(workbook), Structure.Of(reloaded));
        Assert.Equal(workbook.Sheets.Select(s => (s.Name, s.Visibility, s.TabColor)), reloaded.Sheets.Select(s => (s.Name, s.Visibility, s.TabColor)));
        return saved;
    }

    private static List<string> Entries(byte[] file)
    {
        using var zip = new ZipArchive(new MemoryStream(file));
        return [.. zip.Entries.Select(e => e.FullName)];
    }

    [Fact]
    public void A_renamed_sheet_is_renamed_in_the_file_and_in_every_formula()
    {
        var (original, workbook) = Load();

        workbook["B"].Name = "Data 2";
        var saved = SaveChecked(original, workbook);

        var book = Part(saved, "xl/workbook.xml");
        Assert.Contains("<sheet name=\"Data 2\" sheetId=\"2\"", book);
        Assert.Contains(">'Data 2'!$A$1<", book);
        Assert.Contains("name=\"_xlnm.Print_Area\" localSheetId=\"1\">'Data 2'!$A$1:$A$3<", book);
        Assert.Contains("<f>'Data 2'!A1*2</f>", Part(saved, "xl/worksheets/sheet1.xml"));
        Assert.Contains("<c:f>'Data 2'!$A$1:$A$3</c:f>", Part(saved, "xl/charts/chart1.xml"));
        var reloaded = Reload(saved);
        Assert.Equal("A1:A3", reloaded["Data 2"].PrintArea);
        Assert.Equal(20, reloaded["A"]["A1"].Value.AsNumber());
    }

    [Fact]
    public void A_removed_sheet_goes_with_its_parts_and_references_to_it_break()
    {
        var (original, workbook) = Load();

        workbook.RemoveSheet(workbook["B"]);
        var saved = SaveChecked(original, workbook);

        var entries = Entries(saved);
        foreach (var gone in new[] { "xl/worksheets/sheet2.xml", "xl/worksheets/_rels/sheet2.xml.rels", "xl/comments1.xml", "xl/drawings/vmlDrawing1.vml",
                     "xl/tables/table1.xml", "xl/drawings/drawing1.xml", "xl/charts/chart1.xml", "xl/calcChain.xml" })
            Assert.DoesNotContain(gone, entries);
        var types = Part(saved, "[Content_Types].xml");
        Assert.DoesNotContain("sheet2.xml", types);
        Assert.DoesNotContain("table1.xml", types);
        Assert.DoesNotContain("sheet2.xml", Part(saved, "xl/_rels/workbook.xml.rels"));

        var book = Part(saved, "xl/workbook.xml");
        Assert.DoesNotContain("name=\"B\"", book);
        Assert.Contains(">#REF!$A$1<", book);
        Assert.DoesNotContain("Print_Area", book);
        // C moved from the third place to the second: its local name follows, and the window shows it.
        Assert.Contains("name=\"Local\" localSheetId=\"1\"", book);
        Assert.Contains("<workbookView activeTab=\"1\"", book);
        Assert.Contains("<f>#REF!A1*2</f>", Part(saved, "xl/worksheets/sheet1.xml"));
        Assert.Contains("<f>SUM(A:C!A1)</f>", Part(saved, "xl/worksheets/sheet1.xml"));

        var reloaded = Reload(saved);
        reloaded.Recalculate();
        Assert.Equal(CellValue.Error(ErrorKind.Ref), reloaded["A"]["A1"].Value);
        Assert.Equal(CellValue.Error(ErrorKind.Ref), reloaded["A"]["B1"].Value);   // A!A1 is #REF! now
    }

    [Fact]
    public void Moved_sheets_are_saved_in_their_new_order_with_their_names_and_window()
    {
        var (original, workbook) = Load();

        workbook.MoveSheet(workbook["C"], 0);
        var saved = SaveChecked(original, workbook);

        var book = Part(saved, "xl/workbook.xml");
        Assert.True(book.IndexOf("name=\"C\"", StringComparison.Ordinal) < book.IndexOf("name=\"A\"", StringComparison.Ordinal));
        Assert.Contains("name=\"Local\" localSheetId=\"0\"", book);
        Assert.Contains("name=\"_xlnm.Print_Area\" localSheetId=\"2\"", book);
        Assert.Contains("<workbookView activeTab=\"2\"", book);
        Assert.Equal(["C", "A", "B"], Reload(saved).Sheets.Select(s => s.Name));
    }

    [Fact]
    public void Hidden_sheets_and_tab_colours_are_saved()
    {
        var (original, workbook) = Load();

        workbook["B"].Visibility = SheetVisibility.Hidden;
        workbook["C"].Visibility = SheetVisibility.VeryHidden;
        workbook["A"].TabColor = CellColor.FromRgb(0x00B050);
        var saved = SaveChecked(original, workbook);

        var book = Part(saved, "xl/workbook.xml");
        Assert.Contains("<sheet name=\"B\" sheetId=\"2\" r:id=\"rIdS2\" state=\"hidden\"", book);
        Assert.Contains("state=\"veryHidden\"", book);
        // The window showed B, which is hidden now: it shows the next visible sheet, A.
        Assert.DoesNotContain("activeTab", book);
        Assert.Contains("<sheetPr><tabColor rgb=\"FF00B050\" /></sheetPr>", Part(saved, "xl/worksheets/sheet1.xml"));
    }

    [Fact]
    public void A_tab_colour_from_the_theme_is_read_as_rgb()
    {
        var workbook = XlsxWriterTests.Load(new TestXlsx().RawSheet("S", """<sheetPr><tabColor indexed="10"/></sheetPr><sheetData/>"""));

        Assert.Equal(CellColor.FromRgb(0xFF0000), workbook["S"].TabColor);
    }

    [Fact]
    public void A_clone_of_a_workbook_without_a_sheet_saves_without_it()
    {
        var (original, workbook) = Load();
        workbook.RemoveSheet(workbook["C"]);

        var saved = SaveChecked(original, workbook.Clone());

        Assert.DoesNotContain("xl/worksheets/sheet3.xml", Entries(saved));
        Assert.Equal(["A", "B"], Reload(saved).Sheets.Select(s => s.Name));
    }

    [Fact]
    public void A_table_named_as_one_on_a_removed_sheet_is_saved()
    {
        var (original, workbook) = Load();
        workbook.RemoveSheet(workbook["B"]);
        workbook["C"]["B1"].Value = "Units";
        workbook["C"]["B2"].Value = 1;
        workbook["C"].AddTable("Sales", "B1:B2");

        var saved = SaveChecked(original, workbook);

        Assert.Equal("B1:B2", Assert.Single(Reload(saved).Tables).Range);
    }
}
