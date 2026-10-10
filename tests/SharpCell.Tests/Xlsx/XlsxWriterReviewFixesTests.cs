using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

/// <summary>Cases found by the whole-branch review of the writer.</summary>
public class XlsxWriterReviewFixesTests
{
    private static readonly XlsxWriteOptions Keep = new() { KeepUncalculated = true };

    private static (byte[] Original, Workbook Workbook) Load(TestXlsx file)
    {
        var original = file.Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        return (original, workbook);
    }

    [Fact]
    public void Rows_cannot_move_on_a_sheet_whose_formulas_SharpCell_cannot_read()
    {
        var (_, workbook) = Load(new TestXlsx().Sheet("S",
            "<row r=\"5\"><c r=\"A5\"><v>7</v></c><c r=\"B5\"><f>[1]Data!B1+A5</f><v>9</v></c></row>"));

        Assert.Throws<NotSupportedException>(() => workbook["S"].InsertRows(1));
        Assert.Equal("=[1]Data!B1+A5", workbook["S"]["B5"].Formula);
    }

    [Fact]
    public void A_formula_SharpCell_cannot_read_on_another_sheet_keeps_its_text_and_result()
    {
        var (_, workbook) = Load(new TestXlsx()
            .Sheet("S", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>")
            .Sheet("Links", "<row r=\"1\"><c r=\"A1\"><f>[1]Data!B1*2</f><v>42</v></c></row>"));
        workbook["S"].InsertRows(1);
        workbook.Recalculate();

        var saved = Part(Save(workbook, Keep), "xl/worksheets/sheet2.xml");

        Assert.Contains("<f>[1]Data!B1*2</f><v>42</v>", saved);
    }

    [Fact]
    public void A_shared_group_SharpCell_cannot_read_is_kept_as_the_file_has_it()
    {
        var (original, workbook) = Load(new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><f t=\"shared\" ref=\"A1:A3\" si=\"0\">[1]Data!B1*2</f><v>2</v></c><c r=\"C1\"><f>1+1</f><v>2</v></c></row>"
            + "<row r=\"2\"><c r=\"A2\"><f t=\"shared\" si=\"0\"/><v>4</v></c></row><row r=\"3\"><c r=\"A3\"><f t=\"shared\" si=\"0\"/><v>6</v></c></row>"));
        workbook["S"]["C1"].Formula = "=1+2";
        workbook.Recalculate();

        var saved = Save(workbook, Keep);

        var sheet = Part(saved, "xl/worksheets/sheet1.xml");
        Assert.Contains("<c r=\"A2\"><f t=\"shared\" si=\"0\" /><v>4</v></c>", sheet.Replace("\"/>", "\" />"));
        Assert.Contains("<c r=\"A3\"><f t=\"shared\" si=\"0\" /><v>6</v></c>", sheet.Replace("\"/>", "\" />"));
        Assert.Contains("ref=\"A1:A3\" si=\"0\">[1]Data!B1*2</f><v>2</v>", sheet);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
    }

    [Fact]
    public void Changing_the_formula_holding_a_shared_group_SharpCell_cannot_read_is_refused()
    {
        var (_, workbook) = Load(new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><f t=\"shared\" ref=\"A1:A2\" si=\"0\">[1]Data!B1*2</f><v>2</v></c></row>"
            + "<row r=\"2\"><c r=\"A2\"><f t=\"shared\" si=\"0\"/><v>4</v></c></row>"));
        workbook["S"]["A1"].Formula = "=5";
        workbook.Recalculate();

        Assert.Throws<NotSupportedException>(() => Save(workbook, Keep));
    }

    [Fact]
    public void A_kept_formula_on_an_edited_sheet_moves_and_keeps_its_saved_result()
    {
        var (_, workbook) = Load(new TestXlsx().Sheet("S",
            "<row r=\"5\"><c r=\"A5\"><v>7</v></c><c r=\"B5\"><f>NOSUCHFUNCTION(A5)</f><v>99</v></c></row>"));
        workbook["S"].InsertRows(1);
        workbook.Recalculate();

        var sheet = Part(Save(workbook, Keep), "xl/worksheets/sheet1.xml");

        Assert.Contains("<c r=\"B6\"><f>NOSUCHFUNCTION(A6)</f><v>99</v></c>", sheet);
    }

    [Fact]
    public void Deleting_the_only_range_of_newer_conditional_formats_validations_and_sparklines_leaves_a_valid_file()
    {
        const string ext = """
            <extLst>
            <ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings>
            <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{00000000-0000-0000-0000-000000000001}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule><xm:sqref>A2</xm:sqref></x14:conditionalFormatting>
            </x14:conditionalFormattings></ext>
            <ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
            <x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Other!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>A2</xm:sqref></x14:dataValidation>
            </x14:dataValidations></ext>
            <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
            <x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>S!A2:B2</xm:f><xm:sqref>C2</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>
            </x14:sparklineGroups></ext>
            </extLst>
            """;
        var (original, workbook) = Load(new TestXlsx()
            .RawSheet("S", "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"2\"><c r=\"A2\"><v>2</v></c></row></sheetData>" + ext)
            .Sheet("Other", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"));
        workbook["S"].DeleteRows(2);
        workbook.Recalculate();

        var saved = Save(workbook);

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        Assert.DoesNotContain("extLst", Part(saved, "xl/worksheets/sheet1.xml"));
    }

    [Fact]
    public void Formulas_naming_a_table_whose_rows_were_deleted_get_ref_errors()
    {
        var workbook = new Workbook();
        var s = workbook.AddSheet("S");
        var o = workbook.AddSheet("O");
        s["A1"].Value = "H";
        s["A2"].Value = 1;
        s.AddTable("T", "A1:A3");
        o["A1"].Formula = "=SUM(T[H])";
        workbook.Recalculate();

        s.DeleteRows(1, 3);
        workbook.Recalculate();

        Assert.Equal("=SUM(#REF!)", o["A1"].Formula);
        Assert.Contains("<f>SUM(#REF!)</f>", Part(Save(workbook), "xl/worksheets/sheet2.xml"));
    }

    [Fact]
    public void A_newer_function_SharpCell_does_not_know_keeps_its_prefix()
    {
        var (_, workbook) = Load(new TestXlsx().Sheet("S",
            "<row r=\"2\"><c r=\"A2\"><f>_xlfn.ZZZNEW(1)+B2</f><v>1</v></c></row>"));
        workbook["S"].InsertRows(1);
        workbook.Recalculate();

        Assert.Contains("<f>_xlfn.ZZZNEW(1)+B3</f>", Part(Save(workbook, Keep), "xl/worksheets/sheet1.xml"));
    }

    private static TestXlsx WithPivot() => new TestXlsx()
        .RawSheet("S", "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>")
        .Part("xl/worksheets/_rels/sheet1.xml.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rP\" "
            + "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" Target=\"../pivotTables/pivotTable1.xml\"/></Relationships>")
        .Part("xl/pivotTables/pivotTable1.xml",
            "<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"P\" cacheId=\"1\" dataCaption=\"Values\">"
            + "<location ref=\"A3:B7\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/></pivotTableDefinition>",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml");

    [Fact]
    public void Rows_inserted_into_a_pivot_table_refuse_the_save()
    {
        var (_, workbook) = Load(WithPivot());
        workbook["S"].InsertRows(5, 2);
        workbook.Recalculate();

        Assert.Contains("pivot", Assert.Throws<NotSupportedException>(() => Save(workbook)).Message);
    }

    [Fact]
    public void Rows_inserted_above_a_pivot_table_move_it()
    {
        var (_, workbook) = Load(WithPivot());
        workbook["S"].InsertRows(2, 2);
        workbook.Recalculate();

        Assert.Contains("<location ref=\"A5:B9\"", Part(Save(workbook), "xl/pivotTables/pivotTable1.xml"));
    }
}
