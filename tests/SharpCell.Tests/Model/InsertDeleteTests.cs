using SharpCell.Tests.Xlsx;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Model;

/// <summary>Inserting and deleting rows and columns moves cells and the references to them, as in Excel.</summary>
public class InsertDeleteTests
{
    private static (Workbook W, Worksheet S, Worksheet O) Book()
    {
        var workbook = new Workbook();
        var s = workbook.AddSheet("S");
        var o = workbook.AddSheet("Other");
        for (var row = 1; row <= 5; row++)
            s[row, 1].Value = row * 10;
        s["B1"].Formula = "=SUM(A1:A5)";
        s["B2"].Formula = "=A3*2";
        o["A1"].Formula = "=S!A4+1";
        workbook.DefineName("Fourth", "=S!$A$4");
        workbook.Recalculate();
        return (workbook, s, o);
    }

    [Fact]
    public void Inserting_rows_moves_cells_and_the_references_to_them()
    {
        var (w, s, o) = Book();

        s.InsertRows(3, 2);
        w.Recalculate();

        Assert.Equal([10.0, 20, 0, 0, 30, 40, 50], Enumerable.Range(1, 7).Select(r => s[r, 1].Value.Kind == CellValueKind.Empty ? 0 : s[r, 1].Value.AsNumber()));
        Assert.Equal(("=SUM(A1:A7)", 150.0), (s["B1"].Formula, s["B1"].Value.AsNumber()));
        Assert.Equal(("=A5*2", 60.0), (s["B2"].Formula, s["B2"].Value.AsNumber()));
        Assert.Equal(("=S!A6+1", 41.0), (o["A1"].Formula, o["A1"].Value.AsNumber()));
        Assert.Equal("=S!$A$6", Assert.Single(w.DefinedNames, n => n.Name == "Fourth").Formula);
        s["C1"].Formula = "=Fourth";
        w.Recalculate();
        Assert.Equal(40, s["C1"].Value.AsNumber());
    }

    [Fact]
    public void Deleting_rows_moves_cells_up_and_breaks_references_to_deleted_ones()
    {
        var (w, s, o) = Book();

        s.DeleteRows(3);
        w.Recalculate();

        Assert.Equal([10.0, 20, 40, 50], Enumerable.Range(1, 4).Select(r => s[r, 1].Value.AsNumber()));
        Assert.Equal(CellValue.Empty, s["A5"].Value);
        Assert.Equal(("=SUM(A1:A4)", 120.0), (s["B1"].Formula, s["B1"].Value.AsNumber()));
        Assert.Equal(("=#REF!*2", CellValue.Error(ErrorKind.Ref)), (s["B2"].Formula, s["B2"].Value));
        Assert.Equal(("=S!A3+1", 41.0), (o["A1"].Formula, o["A1"].Value.AsNumber()));
    }

    [Fact]
    public void A_moved_formula_keeps_reading_the_cells_it_read()
    {
        var (w, s, _) = Book();
        s["C4"].Formula = "=A4+A1";

        s.InsertRows(2);
        w.Recalculate();

        Assert.Equal(("=A5+A1", 50.0), (s["C5"].Formula, s["C5"].Value.AsNumber()));
        Assert.Equal(["S!A1", "S!A5"], s["C5"].Precedents.Select(p => p.ToString()));
    }

    [Fact]
    public void Columns_move_the_same_way()
    {
        var (w, s, _) = Book();
        s["D1"].Formula = "=B1+A1";

        s.InsertColumns(2);
        w.Recalculate();
        Assert.Equal(("=C1+A1", 160.0), (s["E1"].Formula, s["E1"].Value.AsNumber()));
        Assert.Equal("=SUM(A1:A5)", s["C1"].Formula);

        s.DeleteColumns(1);
        w.Recalculate();
        Assert.Equal(("=B1+#REF!", CellValue.Error(ErrorKind.Ref)), (s["D1"].Formula, s["D1"].Value));
        Assert.Equal("=SUM(#REF!)", s["B1"].Formula);
    }

    [Fact]
    public void Styles_formats_sizes_hidden_rows_and_merges_move_with_the_rows()
    {
        var (_, s, _) = Book();
        var bold = CellStyle.Default with { Font = CellFont.Default with { Bold = true } };
        s["A4"].Style = bold;
        s["A4"].NumberFormat = "0.00";
        s.SetRowHeight(4, 30);
        s.SetRowHidden(5, true);
        s.Merge("C4:D5");
        s.Merge("C1:D1");
        s.SetColumnWidth(1, 20);

        s.InsertRows(2, 3);

        Assert.Equal((bold, "0.00"), (s["A7"].Style, s["A7"].NumberFormat));
        Assert.Equal(CellStyle.Default, s["A4"].Style);
        Assert.Equal((30.0, (double?)null), (s.RowHeight(7)!.Value, s.RowHeight(4)));
        Assert.True(s.IsRowHidden(8));
        Assert.False(s.IsRowHidden(5));
        Assert.Equal(["C7:D8", "C1:D1"], s.MergedAreas.Select(m => m.Address));
        Assert.Equal(20, s.ColumnWidth(1));
    }

    [Fact]
    public void Rows_inserted_inside_a_merge_or_table_grow_it()
    {
        var (w, s, _) = Book();
        s["F1"].Value = "Units";
        s["F2"].Value = 1;
        s["F3"].Value = 2;
        s.AddTable("Sales", "F1:F3");
        s["G1"].Formula = "=SUM(Sales[Units])";
        s.Merge("H1:H3");

        s.InsertRows(3);
        s["F3"].Value = 4;
        w.Recalculate();

        Assert.Equal("F1:F4", w.Tables.Single().Range);
        Assert.Equal("H1:H4", s.MergedAreas.Single().Address);
        Assert.Equal(7, s["G1"].Value.AsNumber());
    }

    [Fact]
    public void Deleting_all_rows_of_a_table_removes_it()
    {
        var (w, s, _) = Book();
        s["F7"].Value = "Units";
        s["F8"].Value = 1;
        s.AddTable("Sales", "F7:F8");
        s["G1"].Formula = "=SUM(Sales[Units])";

        s.DeleteRows(7, 2);
        w.Recalculate();

        Assert.Empty(w.Tables);
        Assert.Equal(CellValue.Error(ErrorKind.Ref), s["G1"].Value);
    }

    [Fact]
    public void Deleting_a_table_header_row_is_refused()
    {
        var (w, s, _) = Book();
        s["F7"].Value = "Units";
        s["F8"].Value = 1;
        s["F9"].Value = 2;
        s.AddTable("Sales", "F7:F9");

        Assert.Throws<InvalidOperationException>(() => s.DeleteRows(7));
        Assert.Equal("F7:F9", w.Tables.Single().Range);
        Assert.Equal(10, s["A1"].Value.AsNumber());
    }

    [Fact]
    public void Inserting_columns_through_a_table_is_not_supported_yet()
    {
        var (w, s, _) = Book();
        s["F7"].Value = "A";
        s["G7"].Value = "B";
        s["F8"].Value = 1;
        s.AddTable("Sales", "F7:G8");

        Assert.Throws<NotSupportedException>(() => s.InsertColumns(7));
        Assert.Equal("F7:G8", w.Tables.Single().Range);
    }

    [Fact]
    public void A_spill_spills_again_from_its_new_place()
    {
        var (w, s, _) = Book();
        s["E1"].Formula = "=SEQUENCE(3)";
        w.Recalculate();

        s.InsertRows(1);
        w.Recalculate();

        Assert.Equal([1.0, 2, 3], new[] { "E2", "E3", "E4" }.Select(a => s[a].Value.AsNumber()));
        Assert.Equal(CellValue.Empty, s["E1"].Value);
    }

    [Fact]
    public void Deleting_part_of_an_array_formula_is_refused()
    {
        var w = XlsxReader.Load(new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"2\"><c r=\"B2\"><f t=\"array\" ref=\"B2:B3\">A1*2</f><v>2</v></c></row>").Build());
        w.Recalculate();
        var s = w["S"];

        Assert.Throws<InvalidOperationException>(() => s.DeleteRows(3));
        Assert.Throws<InvalidOperationException>(() => s.InsertRows(3));
        Assert.Equal(2, s["B3"].Value.AsNumber());

        s.InsertRows(1);
        w.Recalculate();
        Assert.Equal((2.0, 2.0), (s["B3"].Value.AsNumber(), s["B4"].Value.AsNumber()));
        Assert.Throws<InvalidOperationException>(() => s["B4"].Value = 5);   // still one array
    }

    [Fact]
    public void Content_pushed_past_the_last_row_is_refused()
    {
        var (_, s, _) = Book();
        s[CellAddress.MaxRow, 1].Value = 1;

        Assert.Throws<InvalidOperationException>(() => s.InsertRows(2));
        Assert.Equal(20, s["A2"].Value.AsNumber());
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    [InlineData(CellAddress.MaxRow, 2)]
    public void Positions_outside_the_sheet_are_refused(int row, int count)
    {
        var (_, s, _) = Book();
        Assert.Throws<ArgumentOutOfRangeException>(() => s.DeleteRows(row, count));
        Assert.Throws<ArgumentOutOfRangeException>(() => s.InsertRows(row, count));
    }

    [Fact]
    public void A_clone_keeps_the_edits()
    {
        var (w, s, _) = Book();
        s.InsertRows(2);

        var clone = w.Clone();

        Assert.Equal(s.Edits, clone["S"].Edits);
    }

    [Fact]
    public void A_custom_function_cannot_insert_rows()
    {
        var (w, s, _) = Book();
        w.Functions.Add("GROW", _ => { s.InsertRows(1); return 1; });
        s["C1"].Formula = "=GROW()";
        w.Recalculate();

        Assert.Equal(DiagnosticKind.CustomFunctionFailure, Assert.Single(w.Diagnostics).Kind);
        Assert.Equal(10, s["A1"].Value.AsNumber());
    }
}
