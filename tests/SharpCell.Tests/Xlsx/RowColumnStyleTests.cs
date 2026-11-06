using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>Styles of whole rows and columns: cell over row over column, read, saved and moved as Excel has them.</summary>
public class RowColumnStyleTests
{
    private static readonly CellStyle Yellow = CellStyle.Default with { Fill = CellColor.FromRgb(0xFFFF00) };
    private static readonly CellStyle Bold = CellStyle.Default with { Font = CellFont.Default with { Bold = true } };

    private const string Styles = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
          <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellXfs count="3"><xf fontId="0" fillId="0" borderId="0"/><xf fontId="1" fillId="0" borderId="0" applyFont="1"/><xf fontId="0" fillId="2" borderId="0" applyFill="1"/></cellXfs>
        </styleSheet>
        """;

    private static (byte[] Original, Workbook Workbook) Load()
    {
        var original = new TestXlsx { Styles = Styles }.RawSheet("S", """
            <cols><col min="2" max="3" width="9" style="1" customWidth="1"/></cols>
            <sheetData>
              <row r="2" s="2" customFormat="1"><c r="A2" s="2"><v>1</v></c><c r="B2"><v>2</v></c></row>
              <row r="4" s="1"><c r="A4"><v>3</v></c></row>
            </sheetData>
            """).Build().ToArray();
        return (original, XlsxReader.Load(new MemoryStream(original)));
    }

    [Fact]
    public void A_cell_without_a_style_shows_its_rows_then_its_columns()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet.SetColumnStyle(2, Bold);
        sheet.SetRowStyle(3, Yellow);

        Assert.Equal(Bold, sheet["B1"].Style);
        Assert.Equal(Yellow, sheet["B3"].Style);
        Assert.Equal(Yellow, sheet["A3"].Style);
        Assert.Equal(CellStyle.Default, sheet["A1"].Style);
        sheet["B3"].Style = CellStyle.Default;
        Assert.Equal(CellStyle.Default, sheet["B3"].Style);   // its own, even the default, comes first
        sheet.SetRowStyle(3, null);
        Assert.Equal(Bold, sheet["B4"].Style);
        Assert.Null(sheet.RowStyle(3));
    }

    [Fact]
    public void Row_and_column_styles_move_with_inserted_rows_and_columns_and_are_cloned()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet.SetRowStyle(3, Yellow);
        sheet.SetColumnStyle(2, Bold);

        sheet.InsertRows(1, 2);
        sheet.InsertColumns(1);

        Assert.Equal((Yellow, Bold), (sheet.RowStyle(5), sheet.ColumnStyle(3)));
        Assert.Null(sheet.RowStyle(3));
        Assert.Equal(Yellow, workbook.Clone()["S"].RowStyle(5));
    }

    [Fact]
    public void Styles_of_rows_and_columns_are_read_with_customFormat_only()
    {
        var (_, workbook) = Load();
        var sheet = workbook["S"];

        Assert.Equal(Yellow, sheet.RowStyle(2));
        Assert.Null(sheet.RowStyle(4));   // s without customFormat
        Assert.Equal((Bold, Bold, (CellStyle?)null), (sheet.ColumnStyle(2), sheet.ColumnStyle(3), sheet.ColumnStyle(4)));
        Assert.Equal(CellStyle.Default, sheet["B2"].Style);   // its own Normal, not the row's
        Assert.Equal(Yellow, sheet["D2"].Style);
        Assert.Equal(Bold, sheet["C9"].Style);
    }

    [Fact]
    public void Unchanged_rows_and_columns_save_as_the_file_has_them()
    {
        var (original, workbook) = Load();
        workbook["S"]["A2"].Value = 5;

        var part = XlsxWriterTests.Part(XlsxWriterTests.Save(workbook), "xl/worksheets/sheet1.xml");

        Assert.Contains("<row r=\"2\" s=\"2\" customFormat=\"1\">", part);
        Assert.Contains("<col min=\"2\" max=\"3\" width=\"9\" style=\"1\" customWidth=\"1\" />", part);
    }

    [Fact]
    public void Changed_row_and_column_styles_are_saved_and_read_back()
    {
        var (original, workbook) = Load();
        var sheet = workbook["S"];
        sheet.SetRowStyle(2, null);
        sheet.SetRowStyle(6, Bold);
        sheet.SetColumnStyle(3, null);
        sheet.SetColumnStyle(5, Yellow);
        sheet["E7"].Value = "new";   // in a styled column: written with the column's style

        var saved = XlsxWriterTests.Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var reloaded = XlsxWriterTests.Reload(saved)["S"];

        Assert.Equal(((CellStyle?)null, Bold, Bold, (CellStyle?)null, Yellow), (reloaded.RowStyle(2), reloaded.RowStyle(6), reloaded.ColumnStyle(2), reloaded.ColumnStyle(3), reloaded.ColumnStyle(5)));
        Assert.Equal(Yellow, reloaded["E7"].Style);
        Assert.Equal(Bold, reloaded["E6"].Style);   // row before column
    }

    [Fact]
    public void A_workbook_built_in_code_saves_styled_rows_without_cells()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet.SetRowStyle(3, Yellow);
        sheet.SetColumnStyle(1, Bold);
        sheet["B3"].Value = 1;

        var saved = XlsxWriterTests.Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        var reloaded = XlsxWriterTests.Reload(saved)["S"];
        Assert.Equal((Yellow, Bold, Yellow), (reloaded.RowStyle(3), reloaded.ColumnStyle(1), reloaded["B3"].Style));
    }
}
