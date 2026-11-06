using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>Outlines are read from files and saved back with the levels Excel needs to show its buttons.</summary>
public class XlsxOutlineTests
{
    private static (byte[] Original, Workbook Workbook) Load()
    {
        var original = new TestXlsx().RawSheet("S", """
            <sheetPr><outlinePr summaryBelow="0"/></sheetPr>
            <sheetFormatPr defaultRowHeight="15" outlineLevelRow="2" outlineLevelCol="1"/>
            <cols><col min="2" max="3" width="9" hidden="1" outlineLevel="1" customWidth="1"/><col min="4" max="4" width="9" collapsed="1" customWidth="1"/></cols>
            <sheetData>
              <row r="1" collapsed="1"><c r="A1"><v>1</v></c></row>
              <row r="2" outlineLevel="1" hidden="1"><c r="A2"><v>2</v></c></row>
              <row r="3" outlineLevel="2" hidden="1"><c r="A3"><v>3</v></c></row>
              <row r="4" outlineLevel="1" hidden="1"><c r="A4"><v>4</v></c></row>
            </sheetData>
            """).Build().ToArray();
        return (original, XlsxReader.Load(new MemoryStream(original)));
    }

    [Fact]
    public void Outlines_are_read()
    {
        var (_, workbook) = Load();
        var sheet = workbook["S"];

        Assert.Equal([0, 1, 2, 1, 0], Enumerable.Range(1, 5).Select(sheet.RowOutlineLevel));
        Assert.False(sheet.OutlineSummaryBelow);
        Assert.True(sheet.OutlineSummaryRight);
        Assert.True(sheet.IsRowGroupCollapsed(2));   // summary row 1, above
        Assert.Equal((1, 1, 0), (sheet.ColumnOutlineLevel(2), sheet.ColumnOutlineLevel(3), sheet.ColumnOutlineLevel(4)));
        Assert.True(sheet.IsColumnGroupCollapsed(3));
    }

    [Fact]
    public void Unchanged_outlines_save_as_the_file_has_them()
    {
        var (original, workbook) = Load();
        workbook["S"]["A1"].Value = 5;

        var part = XlsxWriterTests.Part(XlsxWriterTests.Save(workbook), "xl/worksheets/sheet1.xml");

        Assert.Contains("<row r=\"3\" outlineLevel=\"2\" hidden=\"1\">", part);
        Assert.Contains("<sheetFormatPr defaultRowHeight=\"15\" outlineLevelRow=\"2\" outlineLevelCol=\"1\" />", part);
        Assert.Contains("<outlinePr summaryBelow=\"0\" />", part);
    }

    [Fact]
    public void Changed_outlines_are_saved_and_read_back()
    {
        var (original, workbook) = Load();
        var sheet = workbook["S"];
        sheet.ExpandRowGroup(2);
        sheet.OutlineSummaryBelow = true;
        sheet.OutlineSummaryRight = false;
        sheet.UngroupRows(3, 3);
        sheet.GroupRows(7, 9);
        sheet.GroupRows(8, 8);
        sheet.GroupRows(8, 8);
        sheet.CollapseRowGroup(8);
        sheet.GroupColumns(6, 7);

        var saved = XlsxWriterTests.Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var part = XlsxWriterTests.Part(saved, "xl/worksheets/sheet1.xml");
        var reloaded = XlsxWriterTests.Reload(saved)["S"];

        Assert.Contains("outlineLevelRow=\"3\"", part);
        Assert.Contains("<outlinePr summaryRight=\"0\" />", part);
        Assert.Equal(Enumerable.Range(1, 10).Select(sheet.RowOutlineLevel), Enumerable.Range(1, 10).Select(reloaded.RowOutlineLevel));
        Assert.Equal(Enumerable.Range(1, 10).Select(sheet.IsRowHidden), Enumerable.Range(1, 10).Select(reloaded.IsRowHidden));
        Assert.True(reloaded.IsRowGroupCollapsed(8));
        Assert.False(reloaded.IsRowGroupCollapsed(2));
        Assert.Equal((1, 1), (reloaded.ColumnOutlineLevel(6), reloaded.ColumnOutlineLevel(7)));
        Assert.False(reloaded.OutlineSummaryRight);
    }

    [Fact]
    public void Removing_every_group_takes_the_levels_out_of_the_sheet_format()
    {
        var (original, workbook) = Load();
        var sheet = workbook["S"];
        sheet.UngroupRows(1, 10);
        sheet.UngroupRows(1, 10);
        sheet.UngroupColumns(1, 10);

        var part = XlsxWriterTests.Part(XlsxWriterTests.Save(workbook), "xl/worksheets/sheet1.xml");

        Assert.DoesNotContain("outlineLevel", part);
    }

    [Fact]
    public void A_workbook_built_in_code_saves_its_outline()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        for (var row = 1; row <= 6; row++)
            sheet[$"A{row}"].Value = row;
        sheet.GroupRows(2, 5);
        sheet.GroupRows(3, 4);
        sheet.CollapseRowGroup(3);
        sheet.GroupColumns(2, 4);

        var saved = XlsxWriterTests.Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        var reloaded = XlsxWriterTests.Reload(saved)["S"];
        Assert.Equal([0, 1, 2, 2, 1, 0], Enumerable.Range(1, 6).Select(reloaded.RowOutlineLevel));
        Assert.Equal([false, false, true, true, false, false], Enumerable.Range(1, 6).Select(reloaded.IsRowHidden));
        Assert.True(reloaded.IsRowGroupCollapsed(3));
        Assert.Equal(1, reloaded.ColumnOutlineLevel(3));
        Assert.Contains("outlineLevelRow=\"2\" outlineLevelCol=\"1\"", XlsxWriterTests.Part(saved, "xl/worksheets/sheet1.xml"));
    }
}
