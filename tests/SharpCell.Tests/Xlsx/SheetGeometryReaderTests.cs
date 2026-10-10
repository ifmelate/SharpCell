using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class SheetGeometryReaderTests
{
    private static Worksheet Load(string content) =>
        XlsxReader.Load(new TestXlsx().RawSheet("S", content).Build())["S"];

    [Fact]
    public void Default_sizes_come_from_sheetFormatPr()
    {
        var sheet = Load("<sheetFormatPr defaultColWidth=\"10.5\" defaultRowHeight=\"14.4\"/><sheetData/>");

        Assert.Equal(10.5, sheet.DefaultColumnWidth);
        Assert.Equal(14.4, sheet.DefaultRowHeight);
    }

    [Fact]
    public void Without_sheetFormatPr_the_sizes_are_unset()
    {
        var sheet = Load("<sheetData/>");

        Assert.Null(sheet.DefaultColumnWidth);
        Assert.Null(sheet.DefaultRowHeight);
    }

    [Fact]
    public void Column_widths_and_hidden_columns_cover_each_col_range()
    {
        var sheet = Load("<cols><col min=\"2\" max=\"3\" width=\"20.7109375\" customWidth=\"1\"/><col min=\"5\" max=\"5\" width=\"0\" hidden=\"1\"/></cols><sheetData/>");

        Assert.Null(sheet.ColumnWidth(1));
        Assert.Equal(20.7109375, sheet.ColumnWidth(2));
        Assert.Equal(20.7109375, sheet.ColumnWidth(3));
        Assert.Null(sheet.ColumnWidth(4));
        Assert.True(sheet.IsColumnHidden(5));
        Assert.False(sheet.IsColumnHidden(4));
    }

    [Fact]
    public void A_col_over_the_rest_of_the_sheet_is_read()
    {
        var sheet = Load("<cols><col min=\"4\" max=\"16384\" width=\"12\"/></cols><sheetData/>");

        Assert.Equal(12, sheet.ColumnWidth(16384));
        Assert.Null(sheet.ColumnWidth(3));
    }

    [Fact]
    public void Row_heights_come_from_ht_with_or_without_customHeight()
    {
        // Excel writes ht for rows grown to fit a larger font too, and shows them at that height.
        var sheet = Load("<sheetData><row r=\"2\" ht=\"30\" customHeight=\"1\"/><row r=\"3\" ht=\"18.75\"><c r=\"A3\"><v>1</v></c></row><row r=\"4\"/></sheetData>");

        Assert.Equal(30, sheet.RowHeight(2));
        Assert.Equal(18.75, sheet.RowHeight(3));
        Assert.Null(sheet.RowHeight(4));
    }

    [Fact]
    public void Merged_cells_are_read_in_order()
    {
        var sheet = Load("<sheetData/><mergeCells count=\"2\"><mergeCell ref=\"A1:C1\"/><mergeCell ref=\"B3:B5\"/></mergeCells>");

        Assert.Equal(["A1:C1", "B3:B5"], sheet.MergedAreas.Select(a => a.Address));
    }

    [Fact]
    public void Broken_geometry_never_stops_a_file()
    {
        var sheet = Load("""
            <sheetFormatPr defaultColWidth="wide" defaultRowHeight="-3"/>
            <cols><col min="9" max="2" width="5"/><col min="1" max="1" width="999"/><col min="0" max="99999" width="5"/></cols>
            <sheetData><row r="1" ht="1e9"/><row r="2"><c r="A2"><v>1</v></c></row></sheetData>
            <mergeCells><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/><mergeCell ref="D1"/><mergeCell ref="nonsense"/></mergeCells>
            """);

        Assert.Null(sheet.DefaultColumnWidth);
        Assert.Null(sheet.DefaultRowHeight);
        Assert.Null(sheet.ColumnWidth(1));
        Assert.Null(sheet.RowHeight(1));
        Assert.Equal("A1:B2", Assert.Single(sheet.MergedAreas).Address);   // overlaps, single cells and garbage are skipped
        Assert.Equal(CellValue.Number(1), sheet["A2"].Value);
    }
}
