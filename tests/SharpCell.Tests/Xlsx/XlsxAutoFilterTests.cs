using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

/// <summary>A sheet's filter is read from the autoFilter element and saved back into it.</summary>
public class XlsxAutoFilterTests
{
    private const string Rows = """
        <row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Units</t></is></c><c r="C1" t="inlineStr"><is><t>Colour</t></is></c></row>
        <row r="2"><c r="A2" t="inlineStr"><is><t>North</t></is></c><c r="B2"><v>10</v></c></row>
        <row r="3" hidden="1"><c r="A3" t="inlineStr"><is><t>South</t></is></c><c r="B3"><v>20</v></c></row>
        <row r="4"><c r="A4" t="inlineStr"><is><t>East</t></is></c><c r="B4"><v>30</v></c></row>
        """;

    // A values filter on A, a top filter on B, a colour filter on C, and the sort Excel saved.
    private static TestXlsx Book() => new TestXlsx().RawSheet("S", $"""
        <sheetPr filterMode="1"/><sheetData>{Rows}</sheetData>
        <autoFilter ref="A1:C4">
          <filterColumn colId="0"><filters blank="1"><filter val="North"/><filter val="East"/></filters></filterColumn>
          <filterColumn colId="1"><top10 val="2" filterVal="20"/></filterColumn>
          <filterColumn colId="2"><colorFilter dxfId="0"/></filterColumn>
          <sortState ref="A2:C4"><sortCondition ref="B2:B4"/></sortState>
        </autoFilter>
        """).Name("_xlnm._FilterDatabase", "S!$A$1:$C$4", " localSheetId=\"0\" hidden=\"1\"");

    [Fact]
    public void A_filter_is_read_with_its_criteria()
    {
        var sheet = Load(Book())["S"];

        var filter = Assert.IsType<AutoFilter>(sheet.AutoFilter);
        Assert.Equal("A1:C4", filter.Range.Address);
        Assert.Equal([1, 2, 3], filter.Columns.Keys);
        Assert.Equal(ColumnFilter.OneOf("North", "East") with { Blanks = true }, filter.Columns[1]);
        Assert.Equal(ColumnFilter.Top(2), filter.Columns[2]);
        Assert.Equal(ColumnFilterType.ByColor, filter.Columns[3].Type);
        Assert.True(sheet.FilterMode);
    }

    [Fact]
    public void Criteria_of_every_kind_read_back_as_they_were_saved()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        for (var row = 1; row <= 4; row++)
        {
            for (var column = 1; column <= 6; column++)
                sheet[row, column].Value = row == 1 ? "H" + column : row * column;
        }

        workbook.Recalculate();
        var filter = sheet.SetAutoFilter("A1:F4");
        ColumnFilter[] filters =
        [
            ColumnFilter.OneOf("2", "8") with { Blanks = true },
            ColumnFilter.OneOf(new FilterDate(DateGrouping.Month, 1900, 1), new FilterDate(DateGrouping.Second, 2026, 10, 17, 9, 30, 15)),
            ColumnFilter.Custom(new FilterCondition(FilterOperator.GreaterThan, "3"), new FilterCondition(FilterOperator.NotEqual, "a*"), matchAll: true),
            ColumnFilter.Bottom(25, percent: true),
            ColumnFilter.Of(DynamicFilter.AboveAverage),
            ColumnFilter.Of(DynamicFilter.ThisQuarter),
        ];
        for (var i = 0; i < filters.Length; i++)
            filter.Filter(i + 1, filters[i]);

        var saved = Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(new TestXlsx().Sheet("S", "").Build().ToArray(), saved);
        var reloaded = Reload(saved)["S"];

        Assert.Equal(filters, reloaded.AutoFilter!.Columns.Values);
        Assert.Equal(Enumerable.Range(2, 3).Select(sheet.IsRowHidden), Enumerable.Range(2, 3).Select(reloaded.IsRowHidden));
        Assert.True(reloaded.FilterMode);
        var part = Part(saved, "xl/worksheets/sheet1.xml");
        Assert.Contains("<dynamicFilter type=\"aboveAverage\" val=\"", part);
        Assert.Contains("<top10 top=\"0\" percent=\"1\" val=\"25\" filterVal=\"", part);
        Assert.Contains("<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"0\" hidden=\"1\">S!$A$1:$F$4</definedName>", Part(saved, "xl/workbook.xml"));
    }

    [Fact]
    public void An_unchanged_filter_leaves_the_file_as_it_was()
    {
        var original = Book().Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();

        Assert.Equal(original, Save(workbook));
    }

    [Fact]
    public void A_changed_filter_keeps_its_colour_column_and_sort_state()
    {
        var workbook = Load(Book());
        var filter = workbook["S"].AutoFilter!;

        filter.Load(1, ColumnFilter.BeginsWith("N"));
        var part = Part(Save(workbook), "xl/worksheets/sheet1.xml");

        Assert.Contains("<filterColumn colId=\"0\"><customFilters><customFilter val=\"N*\" /></customFilters></filterColumn>", part);
        Assert.Contains("<filterColumn colId=\"2\"><colorFilter dxfId=\"0\" /></filterColumn>", part);
        Assert.Contains("<sortState ref=\"A2:C4\">", part);
    }

    [Fact]
    public void A_colour_filter_cannot_be_applied()
    {
        var filter = Load(Book())["S"].AutoFilter!;

        Assert.Throws<NotSupportedException>(() => filter.Apply());
        Assert.Throws<NotSupportedException>(() => filter.Filter(1, ColumnFilter.Blank()));
    }

    [Fact]
    public void A_removed_filter_goes_with_its_name_and_shows_its_rows()
    {
        var workbook = Load(Book());

        workbook["S"].RemoveAutoFilter();
        var saved = Save(workbook);

        Assert.DoesNotContain("autoFilter", Part(saved, "xl/worksheets/sheet1.xml"));
        Assert.DoesNotContain("filterMode", Part(saved, "xl/worksheets/sheet1.xml"));
        Assert.DoesNotContain("_FilterDatabase", Part(saved, "xl/workbook.xml"));
        Assert.False(Reload(saved)["S"].IsRowHidden(3));
    }

    [Fact]
    public void A_filter_moved_by_inserted_rows_is_saved_where_it_went()
    {
        var workbook = Load(Book());

        workbook["S"].InsertRows(1);
        var saved = Save(workbook);

        Assert.Contains("<autoFilter ref=\"A2:C5\">", Part(saved, "xl/worksheets/sheet1.xml"));
        Assert.Contains(">S!$A$2:$C$5<", Part(saved, "xl/workbook.xml"));
        Assert.Equal("A2:C5", Reload(saved)["S"].AutoFilter!.Range.Address);
    }
}
