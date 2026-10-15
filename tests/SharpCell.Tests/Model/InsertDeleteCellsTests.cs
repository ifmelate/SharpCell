using SharpCell.Tests.Xlsx;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Model;

/// <summary>
/// Inserting and deleting cells moves only the cells beside them, as Excel's Insert and Delete with
/// "Shift cells down", "Shift cells right", "Shift cells up" and "Shift cells left" do.
/// </summary>
public class InsertDeleteCellsTests
{
    // A1:C5 hold 11..53 (row * 10 + column); E1 sums column A, E2 column C, E3 all of it.
    private static (Workbook W, Worksheet S) Book()
    {
        var workbook = new Workbook();
        var s = workbook.AddSheet("S");
        for (var row = 1; row <= 5; row++)
        {
            for (var column = 1; column <= 3; column++)
                s[row, column].Value = row * 10 + column;
        }

        s["E1"].Formula = "=SUM(A1:A5)";
        s["E2"].Formula = "=SUM(C1:C5)";
        s["E3"].Formula = "=SUM(A1:C5)";
        s["E4"].Formula = "=A4+C4";
        workbook.Recalculate();
        return (workbook, s);
    }

    private static double?[] Column(Worksheet sheet, int column, int rows) =>
        [.. Enumerable.Range(1, rows).Select(r => sheet[r, column].Value.Kind == CellValueKind.Empty ? (double?)null : sheet[r, column].Value.AsNumber())];

    [Fact]
    public void Shifting_cells_down_moves_only_the_columns_inserted_into()
    {
        var (w, s) = Book();

        s.Range("A2:B3").Insert(InsertShift.Down);
        w.Recalculate();

        Assert.Equal([11, null, null, 21, 31, 41, 51], Column(s, 1, 7));
        Assert.Equal([12, null, null, 22, 32, 42, 52], Column(s, 2, 7));
        Assert.Equal([13, 23, 33, 43, 53, null, null], Column(s, 3, 7));
        Assert.Equal("=SUM(A1:A7)", s["E1"].Formula);   // inside the shifted columns: grows
        Assert.Equal("=SUM(C1:C5)", s["E2"].Formula);   // beside them: stays
        Assert.Equal("=SUM(A1:C5)", s["E3"].Formula);   // across them: stays, as in Excel
        Assert.Equal("=A6+C4", s["E4"].Formula);
        Assert.Equal(41 + 43, s["E4"].Value.AsNumber());
    }

    [Fact]
    public void Shifting_cells_right_moves_only_the_rows_inserted_into()
    {
        var (w, s) = Book();
        s["A7"].Formula = "=SUM(A2:C2)";
        s["A8"].Formula = "=B3";

        s.Range("B2:B3").Insert(InsertShift.Right);
        w.Recalculate();

        Assert.Equal([21.0, 0, 22, 23], Enumerable.Range(1, 4).Select(c => s[2, c].Value.Kind == CellValueKind.Empty ? 0 : s[2, c].Value.AsNumber()));
        Assert.Equal(42, s["B4"].Value.AsNumber());
        Assert.Equal("=SUM(A2:D2)", s["A7"].Formula);
        Assert.Equal("=C3", s["A8"].Formula);
        Assert.Equal(32, s["A8"].Value.AsNumber());
    }

    [Fact]
    public void Deleting_cells_shifts_the_cells_below_up_and_breaks_references_to_them()
    {
        var (w, s) = Book();
        s["F1"].Formula = "=A2";

        s.Range("A2:A3").Delete(DeleteShift.Up);
        w.Recalculate();

        Assert.Equal([11, 41, 51, null, null], Column(s, 1, 5));
        Assert.Equal([12, 22, 32, 42, 52], Column(s, 2, 5));
        Assert.Equal("=SUM(A1:A3)", s["E1"].Formula);
        Assert.Equal(11 + 41 + 51, s["E1"].Value.AsNumber());
        Assert.Equal("=A2+C4", s["E4"].Formula);
        Assert.Equal(("=#REF!", CellValue.Error(ErrorKind.Ref)), (s["F1"].Formula, s["F1"].Value));
    }

    [Fact]
    public void Deleting_cells_can_shift_left()
    {
        var (w, s) = Book();

        s.Range("A1:A2").Delete(DeleteShift.Left);
        w.Recalculate();

        Assert.Equal([12.0, 13], new[] { s["A1"].Value.AsNumber(), s["B1"].Value.AsNumber() });
        Assert.Equal(CellValue.Empty, s["C2"].Value);
        Assert.Equal(31, s["A3"].Value.AsNumber());
        Assert.Equal("=SUM(A1:A5)", s["D1"].Formula);   // E1 is in the shifted rows; part of A1:A5 moved, so it stays
        Assert.Equal("=SUM(A1:C5)", s["E3"].Formula);
    }

    [Fact]
    public void A_moved_formula_keeps_reading_the_same_cells()
    {
        var (w, s) = Book();
        s["B6"].Formula = "=A1+C1";

        s.Range("B2").Insert(InsertShift.Down);
        w.Recalculate();

        Assert.Equal("=A1+C1", s["B7"].Formula);
        Assert.Equal(11 + 13, s["B7"].Value.AsNumber());
    }

    [Fact]
    public void Whole_rows_or_columns_are_inserted_as_rows_or_columns()
    {
        var (w, s) = Book();
        s.SetRowHeight(3, 30);

        s.Range("3:3").Insert(InsertShift.Down);
        s.Range("B:B").Delete(DeleteShift.Left);
        w.Recalculate();

        Assert.Equal(30, s.RowHeight(4));
        Assert.Equal(33, s["B4"].Value.AsNumber());
        Assert.Equal("=SUM(B1:B6)", s["D2"].Formula);
    }

    [Fact]
    public void Styles_formats_and_links_move_with_their_cells()
    {
        var (_, s) = Book();
        s["A3"].Style = s["A3"].Style with { Font = s["A3"].Style.Font with { Bold = true } };
        s["A3"].NumberFormat = "0.00";
        s["C3"].NumberFormat = "0%";

        s.Range("A2:B2").Insert(InsertShift.Down);

        Assert.True(s["A4"].Style.Font.Bold);
        Assert.Equal("0.00", s["A4"].NumberFormat);
        Assert.Equal("0%", s["C3"].NumberFormat);
    }

    [Fact]
    public void Merged_cells_inside_the_shifted_columns_move_and_ones_across_them_refuse()
    {
        var (_, s) = Book();
        s.Merge("A4:B4");

        s.Range("A1:B1").Insert(InsertShift.Down);
        Assert.Equal(["A5:B5"], s.MergedAreas.Select(a => a.Address));

        var error = Assert.Throws<InvalidOperationException>(() => s.Range("A1").Insert(InsertShift.Down));
        Assert.Contains("merged", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(11, s["A2"].Value.AsNumber());   // nothing changed
    }

    [Fact]
    public void Data_validation_across_the_shifted_columns_splits()
    {
        var (_, s) = Book();
        s.AddValidation("A1:C3", DataValidation.List("Yes", "No"));

        s.Range("B1").Insert(InsertShift.Down);

        var areas = Assert.Single(s.Validations.GroupBy(v => v.Validation)).Select(v => v.Range.Address).ToList();
        Assert.Equal(["A1:A3", "B2:B4", "C1:C3"], areas.Order());
    }

    [Fact]
    public void Content_pushed_off_the_sheet_refuses_only_in_the_shifted_columns()
    {
        var (_, s) = Book();
        s[CellAddress.MaxRow, 3].Value = 1;

        s.Range("A1:B1").Insert(InsertShift.Down);
        Assert.Throws<InvalidOperationException>(() => s.Range("C1").Insert(InsertShift.Down));
    }

    [Fact]
    public void A_table_partly_in_the_shifted_columns_refuses()
    {
        var (_, s) = Book();
        s["A10"].Value = "H1";
        s["B10"].Value = "H2";
        s["A11"].Value = 1;
        s["B11"].Value = 2;
        s.AddTable("T", "A10:B11");

        Assert.Throws<InvalidOperationException>(() => s.Range("A2").Insert(InsertShift.Down));
        s.Range("A2:B2").Insert(InsertShift.Down);
        Assert.Equal("A11:B12", s.Workbook.Tables.Single().Range);
    }

    [Fact]
    public void An_array_formula_partly_in_the_shifted_columns_refuses()
    {
        var w = XlsxReader.Load(new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"G1\"><f t=\"array\" ref=\"G1:H2\">A1*2</f><v>2</v></c></row>").Build());
        var s = w["S"];

        Assert.Throws<InvalidOperationException>(() => s.Range("G1").Insert(InsertShift.Down));
        Assert.Throws<InvalidOperationException>(() => s.Range("H1").Delete(DeleteShift.Up));
        s.Range("G1:H1").Insert(InsertShift.Down);
        Assert.Equal(2, s["G2"].Value.AsNumber());
        Assert.Throws<ArgumentException>(() => s.Range("A1").Insert((InsertShift)7));
    }
}
