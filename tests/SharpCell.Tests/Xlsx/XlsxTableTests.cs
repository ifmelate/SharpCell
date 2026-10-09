using System.IO;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class XlsxTableTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string PackageRels = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string TableType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

    private static string Text(string cell, string text) => $"<c r=\"{cell}\" t=\"inlineStr\"><is><t>{text}</t></is></c>";

    private static string Num(string cell, double value) => $"<c r=\"{cell}\"><v>{value}</v></c>";

    // Units in A1, 1/2/4 in A2:A4, row 3 hidden; a formula in C1. Tables are added as parts.
    private static TestXlsx Book(string tables, string afterSheetData = "", string sheetPr = "", string formula = "SUM(Sales[Units])")
    {
        var data = "<sheetData>"
            + $"<row r=\"1\">{Text("A1", "Units")}<c r=\"C1\"><f>{formula}</f><v>0</v></c></row>"
            + $"<row r=\"2\">{Num("A2", 1)}</row>"
            + $"<row r=\"3\" hidden=\"1\">{Num("A3", 2)}</row>"
            + $"<row r=\"4\">{Num("A4", 4)}</row>"
            + "<row r=\"9\" hidden=\"1\"/>"
            + "</sheetData>";
        var xlsx = new TestXlsx().RawSheet("S", sheetPr + data + afterSheetData + tables);
        return xlsx;
    }

    private static TestXlsx WithTable(TestXlsx xlsx, string tableXml)
    {
        return xlsx
            .Part("xl/worksheets/_rels/sheet1.xml.rels",
                $"<Relationships xmlns=\"{PackageRels}\"><Relationship Id=\"rT1\" Type=\"{TableType}\" Target=\"../tables/table1.xml\"/></Relationships>")
            .Part("xl/tables/table1.xml", tableXml);
    }

    private const string TableParts = "<tableParts count=\"1\"><tablePart r:id=\"rT1\"/></tableParts>";

    private static string SalesTable(string extra = "", string autoFilter = "<autoFilter ref=\"A1:A4\"/>") =>
        $"<table xmlns=\"{Main}\" id=\"1\" name=\"Table1\" displayName=\"Sales\" ref=\"A1:A4\"{extra}>{autoFilter}"
        + "<tableColumns count=\"1\"><tableColumn id=\"1\" name=\"Units\"/></tableColumns></table>";

    private static Workbook Load(TestXlsx xlsx)
    {
        using var stream = xlsx.Build();
        return XlsxReader.Load(stream);
    }

    [Fact]
    public void Reads_a_table_and_its_formulas_use_it()
    {
        var wb = Load(WithTable(Book(TableParts), SalesTable()));
        var table = Assert.Single(wb.Tables);
        Assert.Equal("Sales", table.Name);
        Assert.Equal("A1:A4", table.Range);
        Assert.Equal(["Units"], table.Columns);
        Assert.True(table.HasHeaderRow);
        Assert.False(table.HasTotalsRow);

        wb.Recalculate();
        Assert.Equal(CellValue.Number(7), wb["S"]["C1"].Value);
    }

    [Fact]
    public void Reads_header_and_totals_row_counts_and_escaped_column_names()
    {
        var table = "<table xmlns=\"" + Main + "\" id=\"1\" name=\"T\" displayName=\"T\" ref=\"A1:A4\" headerRowCount=\"0\" totalsRowCount=\"1\">"
            + "<tableColumns count=\"1\"><tableColumn id=\"1\" name=\"Unit_x0020_Cost\"/></tableColumns></table>";
        var wb = Load(WithTable(Book(TableParts, formula: "SUM(T[Unit Cost])"), table));
        var t = Assert.Single(wb.Tables);
        Assert.False(t.HasHeaderRow);
        Assert.True(t.HasTotalsRow);
        Assert.Equal(["Unit Cost"], t.Columns);
    }

    [Fact]
    public void Reads_a_single_cell_table_without_header_row()
    {
        var table = "<table xmlns=\"" + Main + "\" id=\"1\" name=\"One\" displayName=\"One\" ref=\"A2\" headerRowCount=\"0\">"
            + "<tableColumns count=\"1\"><tableColumn id=\"1\" name=\"Column1\"/></tableColumns></table>";
        var wb = Load(WithTable(Book(TableParts, formula: "One[Column1]*10"), table));
        Assert.Equal("A2:A2", Assert.Single(wb.Tables).Range);
        wb.Recalculate();
        Assert.Equal(CellValue.Number(10), wb["S"]["C1"].Value);
    }

    [Fact]
    public void Reads_hidden_rows_including_rows_without_cells()
    {
        var s = Load(Book(""))["S"];
        Assert.True(s.IsRowHidden(3));
        Assert.True(s.IsRowHidden(9));
        Assert.False(s.IsRowHidden(2));
        Assert.False(s.FilterMode);
    }

    [Theory]
    [InlineData("<autoFilter ref=\"A1:A4\"><filterColumn colId=\"0\"><filters><filter val=\"1\"/></filters></filterColumn></autoFilter>", true)]
    [InlineData("<autoFilter ref=\"A1:A4\"><filterColumn colId=\"0\"><customFilters><customFilter operator=\"greaterThan\" val=\"0\"/></customFilters></filterColumn></autoFilter>", true)]
    [InlineData("<autoFilter ref=\"A1:A4\"/>", false)]
    [InlineData("<autoFilter ref=\"A1:A4\"><filterColumn colId=\"0\" hiddenButton=\"1\"/></autoFilter>", false)]
    public void Sheet_autoFilter_with_criteria_sets_filter_mode(string autoFilter, bool expected)
    {
        Assert.Equal(expected, Load(Book("", afterSheetData: autoFilter))["S"].FilterMode);
    }

    [Fact]
    public void Table_filter_with_criteria_sets_filter_mode()
    {
        var filter = "<autoFilter ref=\"A1:A4\"><filterColumn colId=\"0\"><filters><filter val=\"1\"/></filters></filterColumn></autoFilter>";
        Assert.True(Load(WithTable(Book(TableParts), SalesTable(autoFilter: filter)))["S"].FilterMode);
    }

    [Fact]
    public void SheetPr_filterMode_sets_filter_mode()
    {
        Assert.True(Load(Book("", sheetPr: "<sheetPr filterMode=\"1\"/>"))["S"].FilterMode);
    }

    [Fact]
    public void AutoFilter_of_a_custom_view_is_not_a_filter()
    {
        var views = "<customSheetViews><customSheetView guid=\"{00000000-0000-0000-0000-000000000001}\">"
            + "<autoFilter ref=\"A1:A4\"><filterColumn colId=\"0\"><filters><filter val=\"1\"/></filters></filterColumn></autoFilter>"
            + "</customSheetView></customSheetViews>";
        Assert.False(Load(Book("", afterSheetData: views))["S"].FilterMode);
    }

    [Fact]
    public void Table_with_more_columns_than_names_is_invalid_data()
    {
        var table = SalesTable().Replace("ref=\"A1:A4\"", "ref=\"A1:B4\"");
        Assert.Throws<InvalidDataException>(() => Load(WithTable(Book(TableParts), table)));
    }

    [Fact]
    public void Header_row_count_other_than_0_or_1_is_invalid_data()
    {
        Assert.Throws<InvalidDataException>(() => Load(WithTable(Book(TableParts), SalesTable(" headerRowCount=\"2\""))));
    }

    [Fact]
    public void Hidden_row_past_the_last_row_is_invalid_data()
    {
        // The second row has no number, so it follows row 1048576 and lies outside the sheet.
        var xlsx = new TestXlsx().RawSheet("S", "<sheetData><row r=\"1048576\"/><row hidden=\"1\"/></sheetData>");
        Assert.Throws<InvalidDataException>(() => Load(xlsx));
    }

    [Fact]
    public void Missing_table_part_is_invalid_data()
    {
        Assert.Throws<InvalidDataException>(() => Load(Book(TableParts)));
    }
}
