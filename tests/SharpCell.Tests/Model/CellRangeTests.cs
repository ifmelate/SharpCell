using System.Linq;

namespace SharpCell.Tests.Model;

public class CellRangeTests
{
    private static (Workbook Workbook, Worksheet Sheet) NewSheet(string name = "S")
    {
        var workbook = new Workbook();
        return (workbook, workbook.AddSheet(name));
    }

    [Fact]
    public void Range_has_its_corners_and_address()
    {
        var (_, sheet) = NewSheet();
        var range = sheet.Range("B2:D5");

        Assert.Same(sheet, range.Worksheet);
        Assert.Equal((2, 2, 5, 4), (range.FirstRow, range.FirstColumn, range.LastRow, range.LastColumn));
        Assert.Equal("B2:D5", range.Address);
        Assert.Equal("S!B2:D5", range.ToString());
    }

    [Fact]
    public void Corners_given_in_any_order_are_normalized()
    {
        var (_, sheet) = NewSheet();
        Assert.Equal("A1:C3", sheet.Range("C3:A1").Address);
    }

    [Fact]
    public void A_single_cell_range_has_a_cell_address()
    {
        var (_, sheet) = NewSheet();
        var range = sheet.Range("B2");
        Assert.Equal("B2", range.Address);
        Assert.Equal((2, 2, 2, 2), (range.FirstRow, range.FirstColumn, range.LastRow, range.LastColumn));
    }

    [Fact]
    public void Whole_columns_and_rows_are_ranges()
    {
        var (_, sheet) = NewSheet();
        var columns = sheet.Range("A:B");
        var rows = sheet.Range("2:3");

        Assert.Equal((1, 1, 1048576, 2), (columns.FirstRow, columns.FirstColumn, columns.LastRow, columns.LastColumn));
        Assert.Equal((2, 1, 3, 16384), (rows.FirstRow, rows.FirstColumn, rows.LastRow, rows.LastColumn));
        Assert.Equal("A1:B1048576", columns.Address);
    }

    [Fact]
    public void A_sheet_name_that_needs_quotes_is_quoted()
    {
        var (_, sheet) = NewSheet("My Sheet");
        Assert.Equal("'My Sheet'!A1:B2", sheet.Range("A1:B2").ToString());
    }

    [Theory]
    [InlineData("")]
    [InlineData("S!A1")]
    [InlineData("A1:B2:C3")]
    [InlineData("Rate")]
    [InlineData("A0")]
    public void Text_that_is_not_a_range_is_rejected(string text)
    {
        var (_, sheet) = NewSheet();
        Assert.Throws<ArgumentException>(() => sheet.Range(text));
    }

    [Fact]
    public void Ranges_are_equal_by_sheet_and_corners()
    {
        var (workbook, sheet) = NewSheet();
        var other = workbook.AddSheet("T");

        Assert.Equal(sheet.Range("A1:B2"), sheet.Range("B2:A1"));
        Assert.Equal(sheet.Range("A1:B2").GetHashCode(), sheet.Range("B2:A1").GetHashCode());
        Assert.NotEqual(sheet.Range("A1:B2"), other.Range("A1:B2"));
        Assert.NotEqual(sheet.Range("A1:B2"), sheet.Range("A1:B3"));
    }

    [Fact]
    public void Cells_lists_non_empty_cells_row_by_row()
    {
        var (workbook, sheet) = NewSheet();
        sheet["C1"].Value = 1;
        sheet["A2"].Formula = "=SEQUENCE(1,2)";
        sheet["E5"].Value = "outside";
        sheet["A1"].Value = "x";
        workbook.Recalculate();

        var addresses = sheet.Range("A1:C3").Cells.Select(c => c.Address).ToArray();

        Assert.Equal(new[] { "A1", "C1", "A2", "B2" }, addresses);
    }

    [Fact]
    public void Cells_of_a_whole_column_visit_only_existing_cells()
    {
        var (_, sheet) = NewSheet();
        sheet["A1000000"].Value = 1;
        sheet["B5"].Value = 2;

        Assert.Equal(new[] { "A1000000" }, sheet.Range("A:A").Cells.Select(c => c.Address));
    }

    [Fact]
    public void Worksheet_cells_lists_every_non_empty_cell()
    {
        var (_, sheet) = NewSheet();
        sheet["B2"].Value = 1;
        sheet["A3"].Value = 2;
        sheet["XFD1048576"].Value = 3;

        Assert.Equal(new[] { "B2", "A3", "XFD1048576" }, sheet.Cells.Select(c => c.Address));
        Assert.Empty(NewSheet().Sheet.Cells);
    }

    [Fact]
    public void Values_of_enumerated_cells_can_be_changed()
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Value = 1;
        sheet["A2"].Value = 2;

        foreach (var cell in sheet.Cells)
            cell.Value = cell.Value.AsNumber() * 10;

        Assert.Equal(10, sheet["A1"].Value.AsNumber());
        Assert.Equal(20, sheet["A2"].Value.AsNumber());
    }

    [Fact]
    public void Adding_or_removing_a_cell_while_enumerating_throws()
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Value = 1;
        sheet["A2"].Value = 2;

        Assert.Throws<InvalidOperationException>(() =>
        {
            foreach (var cell in sheet.Cells)
                sheet["Z9"].Value = 1;
        });
        Assert.Throws<InvalidOperationException>(() =>
        {
            foreach (var cell in sheet.Range("A1:A2").Cells)
                cell.Value = CellValue.Empty;
        });
    }

    [Fact]
    public void UsedRange_bounds_every_non_empty_cell()
    {
        var (workbook, sheet) = NewSheet();
        Assert.Null(sheet.UsedRange);

        sheet["C2"].Value = 1;
        Assert.Equal("C2", sheet.UsedRange!.Address);

        sheet["B5"].Value = 1;
        sheet["E3"].Formula = "=SEQUENCE(1,3)";
        workbook.Recalculate();
        Assert.Equal("B2:G5", sheet.UsedRange!.Address);
    }

    [Fact]
    public void GetValues_returns_every_cell_with_empties()
    {
        var (workbook, sheet) = NewSheet();
        sheet["A1"].Value = 1;
        sheet["B2"].Value = "x";
        sheet["A2"].Formula = "=A1*2";
        workbook.Recalculate();

        var values = sheet.Range("A1:B2").GetValues();

        Assert.Equal(2, values.GetLength(0));
        Assert.Equal(2, values.GetLength(1));
        Assert.Equal(CellValue.Number(1), values[0, 0]);
        Assert.Equal(CellValue.Empty, values[0, 1]);
        Assert.Equal(CellValue.Number(2), values[1, 0]);
        Assert.Equal(CellValue.Text("x"), values[1, 1]);
    }

    [Fact]
    public void SetValues_writes_a_block_and_dependents_follow_after_recalculation()
    {
        var (workbook, sheet) = NewSheet();
        sheet["C1"].Formula = "=SUM(A1:B2)";
        sheet["B1"].Formula = "=99";
        workbook.Recalculate();

        sheet.Range("A1:B2").SetValues(new CellValue[,] { { 1, 2 }, { 3, CellValue.Empty } });
        workbook.Recalculate();

        Assert.Null(sheet["B1"].Formula);
        Assert.Equal(6, sheet["C1"].Value.AsNumber());
        Assert.DoesNotContain(sheet.Cells, c => c.Address == "B2");
    }

    [Fact]
    public void SetValues_checks_everything_before_writing()
    {
        var (_, sheet) = NewSheet();
        var range = sheet.Range("A1:B1");

        Assert.Throws<ArgumentException>(() => range.SetValues(new CellValue[,] { { 1, 2, 3 } }));
        Assert.Throws<ArgumentException>(() => range.SetValues(new CellValue[,] { { 1 }, { 2 } }));
        Assert.Throws<ArgumentException>(() => range.SetValues(new CellValue[,] { { 1, CellValue.Missing } }));
        Assert.Throws<ArgumentNullException>(() => range.SetValues(null!));
        Assert.Empty(sheet.Cells);
    }

    [Fact]
    public void SetValues_refuses_a_part_of_an_array_formula_without_writing_anything()
    {
        using var file = new Xlsx.TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:A2\">{1;2}</f><v>1</v></c></row><row r=\"2\"><c r=\"A2\"><v>2</v></c></row>").Build();
        var sheet = SharpCell.Xlsx.XlsxReader.Load(file)["S"];

        // B1 comes before A2 in the range: nothing may be written when A2 is refused.
        Assert.Throws<InvalidOperationException>(() => sheet.Range("A1:B2").SetValues(new CellValue[,] { { 5, 7 }, { 6, 8 } }));
        Assert.Equal(CellValue.Empty, sheet["B1"].Value);
        Assert.NotNull(sheet["A1"].Formula);
    }
}
