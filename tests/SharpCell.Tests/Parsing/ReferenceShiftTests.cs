using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Parsing;

/// <summary>References in a formula follow inserted and deleted rows and columns, as Excel moves them.</summary>
public class ReferenceShiftTests
{
    private static readonly Workbook Book = MakeBook();

    private static Workbook MakeBook()
    {
        var workbook = new Workbook();
        workbook.AddSheet("S");
        workbook.AddSheet("Other");
        return workbook;
    }

    // The formula in cell `at` of `sheet`, after the edit of sheet S, printed at the cell's new place.
    private static string Shift(string formula, SheetEdit edit, string at = "J10", string sheet = "S", string? movedTo = null)
    {
        CellAddress.TryParse(at, out var origin);
        CellAddress.TryParse(movedTo ?? at, out var newOrigin);
        var node = FormulaParser.Parse("=" + formula, origin);
        var shifted = ReferenceShift.Apply(node, Book[sheet], origin, newOrigin, Book["S"], edit);
        return FormulaPrinter.Print(shifted, newOrigin);
    }

    private static SheetEdit InsertRows(int at, int count) => new(Rows: true, at, count, Delete: false);

    private static SheetEdit DeleteRows(int at, int count) => new(Rows: true, at, count, Delete: true);

    private static SheetEdit InsertColumns(int at, int count) => new(Rows: false, at, count, Delete: false);

    private static SheetEdit DeleteColumns(int at, int count) => new(Rows: false, at, count, Delete: true);

    [Theory]
    [InlineData("A1", "A1")]
    [InlineData("A3", "A5")]
    [InlineData("A1:A5", "A1:A7")]
    [InlineData("A3:A5", "A5:A7")]
    [InlineData("A1:A2", "A1:A2")]
    [InlineData("$A$4", "$A$6")]
    [InlineData("A$4", "A$6")]
    [InlineData("A:A", "A:A")]
    [InlineData("3:4", "5:6")]
    [InlineData("1:3", "1:5")]
    [InlineData("Other!A3", "Other!A3")]
    [InlineData("SUM(A1:A5)*A3", "SUM(A1:A7)*A5")]
    [InlineData("Sales[Units]", "Sales[Units]")]
    [InlineData("Rate", "Rate")]
    public void Inserting_rows_moves_references_at_and_after_the_insertion(string formula, string expected) =>
        Assert.Equal(expected, Shift(formula, InsertRows(3, 2), at: "J1"));

    [Theory]
    [InlineData("A3", "#REF!")]
    [InlineData("SUM(A3)", "SUM(#REF!)")]
    [InlineData("S!A3", "S!#REF!")]
    [InlineData("A1:A5", "A1:A3")]
    [InlineData("A3:A4", "#REF!")]
    [InlineData("A4:A6", "A3:A4")]
    [InlineData("A5", "A3")]
    [InlineData("A2", "A2")]
    [InlineData("3:4", "#REF!")]
    [InlineData("A:A", "A:A")]
    public void Deleting_rows_shrinks_references_and_breaks_deleted_ones(string formula, string expected) =>
        Assert.Equal(expected, Shift(formula, DeleteRows(3, 2), at: "J1"));

    [Fact]
    public void A_formula_on_another_sheet_follows_references_to_the_edited_sheet()
    {
        Assert.Equal("S!A5", Shift("S!A3", InsertRows(3, 2), sheet: "Other"));
        Assert.Equal("A3", Shift("A3", InsertRows(3, 2), sheet: "Other"));
        Assert.Equal("S!A5+s!B5", Shift("'S'!A3+s!B3", InsertRows(3, 2), sheet: "Other"));
    }

    [Fact]
    public void A_moved_formula_keeps_pointing_at_the_same_cells()
    {
        // =A4 in B5; two rows inserted at 1: the formula is now in B7 and reads A6.
        Assert.Equal("A6", Shift("A4", InsertRows(1, 2), at: "B5", movedTo: "B7"));

        // =A1 in B5 (above nothing inserted): A1 stays.
        Assert.Equal("A1", Shift("A1", InsertRows(3, 1), at: "B5", movedTo: "B6"));
    }

    [Theory]
    [InlineData("C1", "E1")]
    [InlineData("A1:C1", "A1:E1")]
    [InlineData("$C$1", "$E$1")]
    [InlineData("A:C", "A:E")]
    [InlineData("C:D", "E:F")]
    [InlineData("1:1", "1:1")]
    public void Inserting_columns_moves_references_the_same_way(string formula, string expected) =>
        Assert.Equal(expected, Shift(formula, InsertColumns(3, 2), at: "A10"));

    [Theory]
    [InlineData("C1", "#REF!")]
    [InlineData("A1:E1", "A1:C1")]
    [InlineData("E1", "C1")]
    [InlineData("C:D", "#REF!")]
    public void Deleting_columns_shrinks_and_breaks_references(string formula, string expected) =>
        Assert.Equal(expected, Shift(formula, DeleteColumns(3, 2), at: "A10"));

    [Fact]
    public void An_unchanged_formula_keeps_its_tree()
    {
        var node = FormulaParser.Parse("=Other!A1+SUM(Other!B:B)", new CellAddress(1, 1));

        Assert.Same(node, ReferenceShift.Apply(node, Book["S"], new CellAddress(1, 1), new CellAddress(1, 1), Book["S"], InsertRows(1, 1)));
    }

    [Fact]
    public void A_reference_pushed_past_the_last_row_is_broken()
    {
        Assert.Equal("#REF!", Shift("A1048576", InsertRows(5, 1), at: "J1"));
        Assert.Equal("A10:A1048576", Shift("A9:A1048576", InsertRows(5, 1), at: "J1"));
    }

    [Theory]
    [InlineData(5, 7, 3, 2, false, 7, 9)]
    [InlineData(5, 7, 6, 2, false, 5, 9)]
    [InlineData(5, 7, 8, 2, false, 5, 7)]
    [InlineData(5, 7, 5, 2, false, 7, 9)]
    [InlineData(5, 7, 6, 1, true, 5, 6)]
    [InlineData(5, 7, 1, 2, true, 3, 5)]
    public void A_span_maps_as_a_range_does(int first, int last, int at, int count, bool delete, int newFirst, int newLast) =>
        Assert.Equal((newFirst, newLast), new SheetEdit(true, at, count, delete).MapSpan(first, last, CellAddress.MaxRow));

    [Fact]
    public void A_span_entirely_deleted_is_gone() =>
        Assert.Null(new SheetEdit(true, 4, 5, true).MapSpan(5, 7, CellAddress.MaxRow));
}
