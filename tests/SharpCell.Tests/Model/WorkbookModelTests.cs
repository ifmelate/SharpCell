using System.Globalization;
using SharpCell;

namespace SharpCell.Tests.Model;

public class WorkbookModelTests
{
    [Fact]
    public void Sheets_are_added_in_order()
    {
        var wb = new Workbook();
        var a = wb.AddSheet("Data");
        var b = wb.AddSheet("Лист 2");
        Assert.Equal([a, b], wb.Sheets);
        Assert.Same(wb, a.Workbook);
        Assert.Equal("Лист 2", b.Name);
        Assert.Same(b, wb["лист 2"]);
        Assert.True(wb.TryGetSheet("DATA", out var found));
        Assert.Same(a, found);
        Assert.False(wb.TryGetSheet("Missing", out _));
        Assert.Throws<KeyNotFoundException>(() => wb["Missing"]);
    }

    [Theory]
    [InlineData("")]
    [InlineData("a:b")]
    [InlineData("a/b")]
    [InlineData("a\\b")]
    [InlineData("a?b")]
    [InlineData("a*b")]
    [InlineData("a[b")]
    [InlineData("a]b")]
    [InlineData("'ab")]
    [InlineData("ab'")]
    [InlineData("1234567890123456789012345678901X")]
    public void Invalid_sheet_names_are_rejected(string name)
    {
        Assert.Throws<ArgumentException>(() => new Workbook().AddSheet(name));
    }

    [Fact]
    public void Duplicate_sheet_names_are_rejected_ignoring_case()
    {
        var wb = new Workbook();
        wb.AddSheet("Sheet1");
        Assert.Throws<ArgumentException>(() => wb.AddSheet("SHEET1"));
    }

    [Fact]
    public void Defaults()
    {
        var wb = new Workbook();
        Assert.Same(CultureInfo.InvariantCulture, wb.Culture);
        Assert.Equal(DateSystem.Date1900, wb.DateSystem);
    }

    [Fact]
    public void Cells_store_values()
    {
        var ws = new Workbook().AddSheet("S");
        ws["B3"].Value = 2;
        Assert.Equal(CellValue.Number(2), ws["b3"].Value);
        Assert.Equal(CellValue.Number(2), ws[3, 2].Value);
        var cell = ws[3, 2];
        Assert.Equal("B3", cell.Address);
        Assert.Equal((3, 2), (cell.Row, cell.Column));
        Assert.Same(ws, cell.Worksheet);
        Assert.Equal(CellValue.Empty, ws["Z99"].Value);
    }

    [Theory]
    [InlineData("A0")]
    [InlineData("XFE1")]
    [InlineData("$A$1")]
    [InlineData("A1:B2")]
    [InlineData("")]
    public void Invalid_addresses_are_rejected(string address)
    {
        var ws = new Workbook().AddSheet("S");
        Assert.Throws<ArgumentException>(() => ws[address]);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    [InlineData(1048577, 1)]
    [InlineData(1, 16385)]
    public void Out_of_sheet_positions_are_rejected(int row, int column)
    {
        var ws = new Workbook().AddSheet("S");
        Assert.Throws<ArgumentOutOfRangeException>(() => ws[row, column]);
    }

    [Fact]
    public void Formulas_are_parsed_on_assignment()
    {
        var ws = new Workbook().AddSheet("S");
        ws["A2"].Formula = "=a1*3";
        Assert.Equal("=a1*3", ws["A2"].Formula);
        ws["A3"].Formula = "SUM(1,2)";
        Assert.Equal("=SUM(1,2)", ws["A3"].Formula);
        Assert.Throws<FormulaParseException>(() => ws["A4"].Formula = "=1+");
        Assert.Null(ws["A4"].Formula);
    }

    [Fact]
    public void Setting_a_value_clears_the_formula_and_vice_versa()
    {
        var ws = new Workbook().AddSheet("S");
        ws["A1"].Formula = "=1+1";
        ws["A1"].Value = "text";
        Assert.Null(ws["A1"].Formula);
        Assert.Equal(CellValue.Text("text"), ws["A1"].Value);

        ws["A1"].Formula = "=1+1";
        ws["A1"].Formula = null;
        Assert.Null(ws["A1"].Formula);
        Assert.Equal(CellValue.Empty, ws["A1"].Value);
    }

    [Fact]
    public void Missing_is_not_a_cell_value()
    {
        var ws = new Workbook().AddSheet("S");
        Assert.Throws<ArgumentException>(() => ws["A1"].Value = CellValue.Missing);
    }

    [Fact]
    public void Storage_is_sparse_and_ordered()
    {
        var ws = new Workbook().AddSheet("S");
        ws["C5"].Value = 1;
        ws["A1"].Value = 2;
        ws["B5"].Value = 3;
        ws["A3"].Value = 4;
        ws["Z1000"].Value = 5;
        ws["A3"].Value = CellValue.Empty;

        var cells = ws.Store.Enumerate(1, 1, 10, 3).Select(c => new CellAddress(c.Row, c.Column).ToString()).ToArray();
        Assert.Equal(["A1", "B5", "C5"], cells);
        Assert.Equal(4, ws.Store.Count);
    }

    [Fact]
    public void Whole_column_enumeration_visits_only_existing_cells()
    {
        var ws = new Workbook().AddSheet("S");
        for (var row = 1; row <= 100; row++)
            ws[row * 1000, 2].Value = row;

        var values = ws.Store.Enumerate(1, 2, CellAddress.MaxRow, 2).Select(c => c.Data.Value.AsNumber()).ToArray();
        Assert.Equal(Enumerable.Range(1, 100).Select(n => (double)n), values);
    }

    [Theory]
    [InlineData("Rate")]
    [InlineData("_total")]
    [InlineData("Налог.2026")]
    [InlineData("XYZ1")]
    public void Valid_names_can_be_defined(string name)
    {
        var wb = new Workbook();
        wb.DefineName(name, "=1");
        Assert.True(wb.Names.TryGet(name.ToUpperInvariant(), null, out _));
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("R1C1")]
    [InlineData("R")]
    [InlineData("TRUE")]
    [InlineData("1x")]
    [InlineData("a b")]
    [InlineData("")]
    [InlineData("Sheet1!x")]
    public void Invalid_names_are_rejected(string name)
    {
        Assert.Throws<ArgumentException>(() => new Workbook().DefineName(name, "=1"));
    }

    [Fact]
    public void Sheet_scoped_names_are_separate_from_workbook_names()
    {
        var wb = new Workbook();
        var ws = wb.AddSheet("S");
        wb.DefineName("Rate", "=1");
        wb.DefineName("rate", "=2", ws);
        Assert.True(wb.Names.TryGet("RATE", null, out var global));
        Assert.True(wb.Names.TryGet("RATE", ws, out var local));
        Assert.Equal("=1", global!.Text);
        Assert.Equal("=2", local!.Text);
    }

    [Fact]
    public void Names_from_another_workbook_sheet_are_rejected()
    {
        var other = new Workbook().AddSheet("S");
        Assert.Throws<ArgumentException>(() => new Workbook().DefineName("x", "=1", other));
    }
}
