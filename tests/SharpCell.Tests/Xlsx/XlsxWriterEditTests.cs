using System.Text.RegularExpressions;
using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

/// <summary>After rows or columns are inserted or deleted, a saved file has everything where the model has it.</summary>
public class XlsxWriterEditTests
{
    private const string Rels = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string TableType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    private const string TableContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

    private const string SheetXml = """
        <dimension ref="A1:D6"/>
        <sheetViews><sheetView workbookViewId="0"><selection activeCell="B5" sqref="B5:C6"/></sheetView></sheetViews>
        <sheetFormatPr defaultRowHeight="15"/>
        <cols><col min="2" max="2" width="20" customWidth="1"/><col min="4" max="4" width="9" hidden="1" customWidth="1"/></cols>
        <sheetData>
          <row r="1" spans="1:3"><c r="A1"><v>1</v></c><c r="B1"><f>SUM(A1:A6)</f><v>21</v></c></row>
          <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B4" si="0">A2*10</f><v>20</v></c></row>
          <row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>30</v></c></row>
          <row r="4" ht="30" customHeight="1"><c r="A4"><v>4</v></c><c r="B4"><f t="shared" si="0"/><v>40</v></c></row>
          <row r="5" hidden="1"><c r="A5"><v>5</v></c><c r="C5"><f t="array" ref="C5:C6">A5:A6*2</f><v>10</v></c></row>
          <row r="6"><c r="A6"><v>6</v></c><c r="C6"><v>12</v></c></row>
        </sheetData>
        <autoFilter ref="A1:B6"/>
        <mergeCells count="1"><mergeCell ref="C2:D3"/></mergeCells>
        <conditionalFormatting sqref="A2:A6 C2"><cfRule type="expression" priority="1"><formula>A2&gt;$A$4</formula></cfRule></conditionalFormatting>
        <dataValidations count="1"><dataValidation type="whole" allowBlank="1" sqref="D5:D6"><formula1>$A$1</formula1><formula2>$A$6</formula2></dataValidation></dataValidations>
        <hyperlinks><hyperlink ref="A4" location="S!A1" display="top"/></hyperlinks>
        <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        <rowBreaks count="1" manualBreakCount="1"><brk id="4" max="16383" man="1"/></rowBreaks>
        <ignoredErrors><ignoredError sqref="A3:A4" numberStoredAsText="1"/></ignoredErrors>
        """;

    private static TestXlsx Book() => new TestXlsx().RawSheet("S", SheetXml)
        .Sheet("Other", "<row r=\"1\"><c r=\"A1\"><f>S!A4*2</f><v>8</v></c></row>")
        .Name("_xlnm.Print_Area", "S!$A$1:$B$6", " localSheetId=\"0\"")
        .Name("_xlnm._FilterDatabase", "S!$A$1:$B$6", " localSheetId=\"0\" hidden=\"1\"");

    private static (string Sheet, string Workbook, Workbook Model, Workbook Reloaded) Edit(Action<Worksheet> change, TestXlsx? file = null)
    {
        var original = (file ?? Book()).Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        change(workbook["S"]);
        workbook.Recalculate();
        var saved = Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var reloaded = Reload(saved);
        foreach (var sheet in workbook.Sheets)
            Assert.Equal(Describe(sheet), Describe(reloaded[sheet.Name]));
        Assert.Equal(Presentation.Of(workbook), Presentation.Of(reloaded));
        Assert.Equal(Structure.Of(workbook), Structure.Of(reloaded));
        return (Part(saved, "xl/worksheets/sheet1.xml"), Part(saved, "xl/workbook.xml"), workbook, reloaded);
    }

    private static List<string> Describe(Worksheet sheet) =>
        [.. sheet.Cells.Select(c => $"{c.Address} {c.Formula} {c.Value} {c.NumberFormat}")];

    [Fact]
    public void Inserted_rows_move_cells_and_everything_that_points_at_them()
    {
        var (sheet, workbook, _, reloaded) = Edit(s => s.InsertRows(3, 2));

        Assert.Equal(8, reloaded["Other"]["A1"].Value.AsNumber());   // =S!A6*2, the old A4
        Assert.Equal("=S!A6*2", reloaded["Other"]["A1"].Formula);
        Assert.Contains("<autoFilter ref=\"A1:B8\"", sheet);
        Assert.Contains("<mergeCell ref=\"C2:D5\"", sheet);
        Assert.Contains("sqref=\"A2:A8 C2\"", sheet);
        Assert.Contains("<formula>A2&gt;$A$6</formula>", sheet);
        Assert.Contains("sqref=\"D7:D8\"", sheet);
        Assert.Contains("<formula2>$A$8</formula2>", sheet);
        Assert.Contains("<hyperlink ref=\"A6\"", sheet);
        Assert.Contains("<brk id=\"6\"", sheet);
        Assert.Contains("<ignoredError sqref=\"A5:A6\"", sheet);
        Assert.Contains("activeCell=\"B7\" sqref=\"B7:C8\"", sheet);
        Assert.Contains(">S!$A$1:$B$8</definedName>", workbook);
        Assert.DoesNotContain("spans=", sheet);
        Assert.DoesNotContain("t=\"shared\"", sheet);
    }

    [Fact]
    public void Deleted_rows_shrink_or_drop_what_pointed_at_them()
    {
        var (sheet, workbook, _, _) = Edit(s => s.DeleteRows(2, 3));

        Assert.Contains("<autoFilter ref=\"A1:B3\"", sheet);
        Assert.DoesNotContain("mergeCell", sheet);
        Assert.Contains("sqref=\"A2:A3\"", sheet);
        Assert.True(sheet.Contains("<formula>A2&gt;#REF!</formula>"), sheet);
        Assert.Contains("sqref=\"D2:D3\"", sheet);
        Assert.DoesNotContain("hyperlink", sheet);
        Assert.DoesNotContain("brk", sheet);
        Assert.DoesNotContain("ignoredError", sheet);
        Assert.Contains("activeCell=\"B2\" sqref=\"B2:C3\"", sheet);
        Assert.Contains(">S!$A$1:$B$3</definedName>", workbook);
        Assert.Contains("<f t=\"array\" ref=\"C2:C3\">A2:A3*2</f>", sheet);
    }

    [Fact]
    public void Inserted_columns_move_widths_and_ranges()
    {
        var (sheet, _, model, _) = Edit(s => s.InsertColumns(2));

        Assert.Equal((20.0, true), (model["S"].ColumnWidth(3)!.Value, model["S"].IsColumnHidden(5)));
        Assert.Contains("<col min=\"3\" max=\"3\" width=\"20\" customWidth=\"1\"", sheet);
        Assert.Contains("sqref=\"A2:A6 D2\"", sheet);
        Assert.Contains("<mergeCell ref=\"D2:E3\"", sheet);
        Assert.Contains("<autoFilter ref=\"A1:C6\"", sheet);
    }

    [Fact]
    public void Deleted_columns_drop_their_cells()
    {
        var (sheet, _, _, reloaded) = Edit(s => s.DeleteColumns(4));

        Assert.DoesNotContain("hidden=\"1\" customWidth", sheet);
        Assert.DoesNotContain("dataValidation ", sheet);
        Assert.Equal(21, reloaded["S"]["B1"].Value.AsNumber());
    }

    [Fact]
    public void A_moved_table_is_saved_where_it_is_and_a_removed_one_goes()
    {
        TestXlsx WithTable() => new TestXlsx().RawSheet("S",
                "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Units</t></is></c></row><row r=\"2\"><c r=\"A2\"><v>1</v></c></row>"
                + "<row r=\"3\"><c r=\"A3\"><v>2</v></c></row></sheetData><tableParts count=\"1\"><tablePart r:id=\"rT1\"/></tableParts>")
            .Part("xl/worksheets/_rels/sheet1.xml.rels",
                $"<Relationships xmlns=\"{Rels}\"><Relationship Id=\"rT1\" Type=\"{TableType}\" Target=\"../tables/table1.xml\"/></Relationships>")
            .Part("xl/tables/table1.xml",
                "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"Sales\" displayName=\"Sales\" ref=\"A1:A3\">"
                + "<autoFilter ref=\"A1:A3\"/><tableColumns count=\"1\"><tableColumn id=\"1\" name=\"Units\"/></tableColumns></table>", TableContentType);

        var original = WithTable().Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        workbook["S"].InsertRows(1);
        workbook.Recalculate();
        var moved = Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, moved);
        Assert.Contains("ref=\"A2:A4\"", Part(moved, "xl/tables/table1.xml"));
        Assert.Equal("A2:A4", Reload(moved).Tables.Single().Range);

        workbook["S"].DeleteRows(1, 4);
        workbook.Recalculate();
        var removed = Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, removed);
        Assert.Empty(Reload(removed).Tables);
        Assert.DoesNotContain("tablePart", Part(removed, "xl/worksheets/sheet1.xml"));
    }

    [Fact]
    public void A_data_table_on_an_edited_sheet_refuses_the_save()
    {
        var workbook = XlsxReader.Load(new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"2\"><c r=\"B2\"><f t=\"dataTable\" ref=\"B2:B3\" dt2D=\"0\" dtr=\"0\" r1=\"A1\"/><v>5</v></c></row>").Build());
        workbook.Recalculate();
        workbook["S"].InsertRows(1);
        workbook.Recalculate();

        Assert.Contains("data table", Assert.Throws<NotSupportedException>(() => Save(workbook)).Message);
    }

    [Fact]
    public void Ranges_in_an_extension_list_move_too()
    {
        const string ext = """
            <extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
            <x14:cfRule type="expression" priority="2" id="{00000000-0000-0000-0000-000000000001}"><xm:f>$A$4&gt;0</xm:f></x14:cfRule>
            <xm:sqref>A3:A4</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>
            """;
        var file = new TestXlsx().RawSheet("S", "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>" + ext);

        var (sheet, _, _, _) = Edit(s => s.InsertRows(1), file);

        Assert.Contains("<xm:sqref>A4:A5</xm:sqref>", sheet);
        Assert.Contains("<xm:f>$A$5&gt;0</xm:f>", sheet);
    }

    public static TheoryData<string, string> CorpusEdits()
    {
        var data = new TheoryData<string, string>();
        foreach (var file in CorpusFiles.All)
        {
            foreach (var edit in new[] { "insert row", "delete row", "insert column", "delete column" })
                data.Add(file, edit);
        }

        return data;
    }

    [Theory]
    [MemberData(nameof(CorpusEdits))]
    public void A_corpus_file_edited_in_each_sheet_saves_and_reads_back_the_same(string file, string edit)
    {
        var original = File.ReadAllBytes(CorpusFiles.PathOf(file));
        var workbook = XlsxReader.Load(new MemoryStream(original));
        foreach (var sheet in workbook.Sheets)
        {
            try
            {
                switch (edit)
                {
                    case "insert row": sheet.InsertRows(2); break;
                    case "delete row": sheet.DeleteRows(2); break;
                    case "insert column": sheet.InsertColumns(2); break;
                    default: sheet.DeleteColumns(2); break;
                }
            }
            catch (Exception ex) when (ex is InvalidOperationException or NotSupportedException)
            {
                // Arrays and tables that cannot be split stay where they are.
            }
        }

        workbook.Recalculate();
        byte[] saved;
        try
        {
            saved = Save(workbook, new XlsxWriteOptions { KeepUncalculated = true });
        }
        catch (NotSupportedException ex) when (ex.Message.Contains("data table"))
        {
            return;
        }

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var reloaded = Reload(saved);
        foreach (var sheet in workbook.Sheets)
            Assert.Equal(Describe(sheet).Where(KeptOut).ToList(), Describe(reloaded[sheet.Name]).Where(KeptOut).ToList());
        Assert.Equal(Presentation.Of(workbook), Presentation.Of(reloaded));
    }

    // Cells SharpCell could not calculate keep the result the file had, which can differ.
    private static bool KeptOut(string description) => !description.Contains("#NAME?", StringComparison.Ordinal);
}
