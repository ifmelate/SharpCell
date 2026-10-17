namespace SharpCell.Tests.Model;

/// <summary>Renaming, removing, moving and hiding sheets: formulas follow, as in Excel.</summary>
public class SheetOperationsTests
{
    private static (Workbook W, Worksheet A, Worksheet B, Worksheet C) Book()
    {
        var workbook = new Workbook();
        var a = workbook.AddSheet("A");
        var b = workbook.AddSheet("B");
        var c = workbook.AddSheet("C");
        a["A1"].Value = 1;
        b["A1"].Value = 10;
        c["A1"].Value = 100;
        a["B1"].Formula = "=B!A1*2";
        a["B2"].Formula = "=SUM(A:C!A1)";
        a["B3"].Formula = "=INDIRECT(\"B!A1\")";
        c["B1"].Formula = "=B!A1+B!$A$1";
        workbook.DefineName("Ten", "=B!$A$1");
        workbook.Recalculate();
        return (workbook, a, b, c);
    }

    [Fact]
    public void Renaming_a_sheet_renames_it_in_every_formula()
    {
        var (w, a, b, c) = Book();
        a.AddValidation("D1", DataValidation.ListFrom("=B!$A$1:$A$3"));
        a.AddConditionalFormat("D1", ConditionalFormat.Expression("=B!$A$1>5", ConditionalStyle.YellowFill));

        b.Name = "Data 2";
        w.Recalculate();

        Assert.Equal("Data 2", b.Name);
        Assert.Same(b, w["data 2"]);
        Assert.False(w.TryGetSheet("B", out _));
        Assert.Equal(("='Data 2'!A1*2", 20.0), (a["B1"].Formula, a["B1"].Value.AsNumber()));
        Assert.Equal(("=SUM(A:C!A1)", 111.0), (a["B2"].Formula, a["B2"].Value.AsNumber()));
        Assert.Equal("='Data 2'!A1+'Data 2'!$A$1", c["B1"].Formula);
        Assert.Equal("='Data 2'!$A$1", Assert.Single(w.DefinedNames).Formula);
        Assert.Equal("='Data 2'!$A$1:$A$3", Assert.Single(a.Validations).Validation.Formula1);
        Assert.Equal("='Data 2'!$A$1>5", Assert.Single(a.ConditionalFormats).Format.Formula1);

        // INDIRECT's text keeps the old name, as in Excel.
        Assert.Equal(("=INDIRECT(\"B!A1\")", CellValue.Error(ErrorKind.Ref)), (a["B3"].Formula, a["B3"].Value));
    }

    [Fact]
    public void Renaming_the_first_or_last_sheet_of_a_3d_reference_renames_that_end()
    {
        var (w, a, _, c) = Book();

        c.Name = "Last";
        a.Name = "First";
        w.Recalculate();

        Assert.Equal(("=SUM(First:Last!A1)", 111.0), (a["B2"].Formula, a["B2"].Value.AsNumber()));
    }

    [Fact]
    public void A_sheet_cannot_take_a_name_another_sheet_has_or_an_invalid_one()
    {
        var (_, a, b, _) = Book();

        Assert.Throws<ArgumentException>(() => b.Name = "a");
        Assert.Throws<ArgumentException>(() => b.Name = "x/y");
        Assert.Throws<ArgumentException>(() => b.Name = "");
        b.Name = "b";
        Assert.Equal("b", b.Name);
        Assert.Equal("=b!A1*2", a["B1"].Formula);
    }

    [Fact]
    public void Removing_a_sheet_breaks_references_to_it_as_Excel_does()
    {
        var (w, a, b, c) = Book();

        w.RemoveSheet(b);
        w.Recalculate();

        Assert.Equal([a, c], w.Sheets);
        Assert.True(b.IsRemoved);
        Assert.Equal(("=#REF!A1*2", CellValue.Error(ErrorKind.Ref)), (a["B1"].Formula, a["B1"].Value));
        Assert.Equal("=#REF!A1+#REF!$A$1", c["B1"].Formula);
        Assert.Equal("=#REF!$A$1", Assert.Single(w.DefinedNames).Formula);
        // A 3D reference keeps the sheets between its ends.
        Assert.Equal(("=SUM(A:C!A1)", 101.0), (a["B2"].Formula, a["B2"].Value.AsNumber()));
    }

    [Fact]
    public void Removing_an_end_of_a_3d_reference_moves_that_end_inwards()
    {
        var (w, a, _, c) = Book();
        c["B2"].Formula = "=SUM(A:C!A1)";

        w.RemoveSheet(a);
        w.Recalculate();

        Assert.Equal(("=SUM('B:C'!A1)", 110.0), (c["B2"].Formula, c["B2"].Value.AsNumber()));
    }

    [Fact]
    public void A_removed_sheet_takes_its_names_and_tables_with_it()
    {
        var (w, a, b, _) = Book();
        b["A2"].Value = "Units";
        b["A3"].Value = 5;
        b.AddTable("Sales", "A2:A3");
        w.DefineName("Local", "=$A$1", b);
        a["C1"].Formula = "=SUM(Sales[Units])";
        a["C2"].Formula = "=B!Local";
        w.Recalculate();

        w.RemoveSheet(b);
        w.Recalculate();

        Assert.Empty(w.Tables);
        Assert.DoesNotContain(w.DefinedNames, name => name.Name == "Local");
        Assert.Equal(("=SUM(#REF!)", CellValue.Error(ErrorKind.Ref)), (a["C1"].Formula, a["C1"].Value));
        Assert.Equal(("=#REF!", CellValue.Error(ErrorKind.Ref)), (a["C2"].Formula, a["C2"].Value));
        // The table's name is free again.
        a.AddTable("Sales", "A1:A2", hasHeaderRow: false);
    }

    [Fact]
    public void A_removed_sheet_can_be_read_but_not_changed()
    {
        var (w, _, b, _) = Book();

        w.RemoveSheet(b);

        Assert.Equal(10, b["A1"].Value.AsNumber());
        Assert.Throws<InvalidOperationException>(() => b["A1"].Value = 2);
        Assert.Throws<InvalidOperationException>(() => b.Name = "Back");
        Assert.Throws<InvalidOperationException>(() => b.InsertRows(1));
        Assert.Throws<ArgumentException>(() => w.RemoveSheet(b));
        Assert.Throws<ArgumentException>(() => w.MoveSheet(b, 0));
    }

    [Fact]
    public void A_new_sheet_with_a_removed_sheets_name_is_not_the_old_one()
    {
        var (w, a, b, _) = Book();

        w.RemoveSheet(b);
        var again = w.AddSheet("B");
        again["A1"].Value = 7;
        w.Recalculate();

        Assert.Equal(("=#REF!A1*2", CellValue.Error(ErrorKind.Ref)), (a["B1"].Formula, a["B1"].Value));
    }

    [Fact]
    public void Moving_a_sheet_changes_which_sheets_a_3d_reference_takes_in()
    {
        var (w, a, b, c) = Book();

        w.MoveSheet(b, 2);
        w.Recalculate();

        Assert.Equal([a, c, b], w.Sheets);
        Assert.Equal(("=SUM(A:C!A1)", 101.0), (a["B2"].Formula, a["B2"].Value.AsNumber()));

        w.MoveSheet(b, 1);
        w.Recalculate();
        Assert.Equal(111, a["B2"].Value.AsNumber());
        Assert.Throws<ArgumentOutOfRangeException>(() => w.MoveSheet(b, 3));
    }

    [Fact]
    public void A_workbook_keeps_one_visible_sheet()
    {
        var (w, a, b, c) = Book();

        b.Visibility = SheetVisibility.Hidden;
        c.Visibility = SheetVisibility.VeryHidden;

        Assert.Throws<InvalidOperationException>(() => a.Visibility = SheetVisibility.Hidden);
        Assert.Throws<InvalidOperationException>(() => w.RemoveSheet(a));
        b.Visibility = SheetVisibility.Visible;
        w.RemoveSheet(a);
        Assert.Equal([b, c], w.Sheets);
    }

    [Fact]
    public void The_only_sheet_can_be_removed()
    {
        var workbook = new Workbook();
        var only = workbook.AddSheet("Only");

        workbook.RemoveSheet(only);

        Assert.Empty(workbook.Sheets);
    }

    [Fact]
    public void A_clone_keeps_visibility_and_tab_colour()
    {
        var (w, _, b, _) = Book();
        b.Visibility = SheetVisibility.Hidden;
        b.TabColor = CellColor.FromRgb(0xFF0000);

        var copy = w.Clone();

        Assert.Equal(SheetVisibility.Hidden, copy["B"].Visibility);
        Assert.Equal(CellColor.FromRgb(0xFF0000), copy["B"].TabColor);
    }

    [Fact]
    public void A_formula_typed_with_a_reference_into_a_removed_sheet_reads_back()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");

        sheet["A1"].Formula = "=SUM(#REF!A1:B2)+1";
        workbook.Recalculate();

        Assert.Equal(("=SUM(#REF!A1:B2)+1", CellValue.Error(ErrorKind.Ref)), (sheet["A1"].Formula, sheet["A1"].Value));
    }
}
