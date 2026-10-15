namespace SharpCell.Tests.Model;

/// <summary>Notes on cells, as Excel's Insert Note makes them.</summary>
public class NoteTests
{
    [Fact]
    public void A_cell_takes_a_note_and_null_removes_it()
    {
        var sheet = new Workbook().AddSheet("S");

        sheet["B2"].Note = new Note("Check this", "Ann") { IsVisible = true };
        sheet["A1"].Note = new Note("First");

        Assert.Equal(("Check this", "Ann", true), (sheet["B2"].Note!.Text, sheet["B2"].Note!.Author, sheet["B2"].Note!.IsVisible));
        Assert.Equal(["A1", "B2"], sheet.Notes.Select(n => n.Cell.Address));
        sheet["B2"].Note = null;
        Assert.Null(sheet["B2"].Note);
        Assert.Equal(CellValue.Empty, sheet["A1"].Value);   // a note is not a value
    }

    [Fact]
    public void Notes_compare_by_text_author_and_visibility()
    {
        Assert.Equal(new Note("a", "x"), new Note("a", "x"));
        Assert.NotEqual(new Note("a", "x"), new Note("a", "x") { IsVisible = true });
        Assert.NotEqual(new Note("a"), new Note("a", "x"));
        Assert.Throws<ArgumentException>(() => new Note(new string('a', 32768)));
    }

    [Fact]
    public void Notes_move_with_their_cells_and_go_with_deleted_ones()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet["A2"].Note = new Note("two");
        sheet["A3"].Note = new Note("three");
        sheet["B3"].Note = new Note("beside");

        sheet.InsertRows(1);
        sheet.DeleteRows(3);
        sheet.Range("A1").Insert(InsertShift.Down);

        Assert.Equal([("B3", "beside"), ("A4", "three")], sheet.Notes.Select(n => (n.Cell.Address, n.Note.Text)));
    }

    [Fact]
    public void Copying_cells_copies_their_notes()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet["A1"].Note = new Note("copied");
        sheet["D2"].Note = new Note("replaced");

        sheet.Range("A1:A2").CopyTo(sheet.Range("D1"));

        Assert.Equal("copied", sheet["D1"].Note!.Text);
        Assert.Null(sheet["D2"].Note);
    }
}
