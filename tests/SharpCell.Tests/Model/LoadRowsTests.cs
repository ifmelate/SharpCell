using System.Globalization;

namespace SharpCell.Tests.Model;

public class LoadRowsTests
{
    private sealed record Sale(string Name, int Units, decimal Price, DateTime Sold, bool Paid);

    private static readonly (string, Func<Sale, object?>)[] SaleColumns =
    [
        ("Name", s => s.Name),
        ("Units", s => s.Units),
        ("Price", s => s.Price),
        ("Sold", s => s.Sold),
        ("Paid", s => s.Paid),
    ];

    private static Sale[] TwoSales() =>
    [
        new("Tea", 3, 4.5m, new DateTime(2026, 3, 14), true),
        new("Coffee", 5, 6m, new DateTime(2026, 3, 15, 9, 30, 0), false),
    ];

    private static Worksheet NewSheet(string culture = "en-US", Workbook? workbook = null)
    {
        workbook ??= new Workbook();
        workbook.Culture = CultureInfo.GetCultureInfo(culture);
        return workbook.AddSheet("S");
    }

    [Fact]
    public void Rows_are_written_under_a_header_and_the_range_is_returned()
    {
        var sheet = NewSheet();
        var range = sheet.LoadRows("B2", TwoSales(), SaleColumns);

        Assert.Equal("B2:F4", range.Address);
        Assert.Equal("Name", sheet["B2"].Text);
        Assert.Equal("Paid", sheet["F2"].Text);
        Assert.Equal("Tea", sheet["B3"].Text);
        Assert.Equal(CellValueKind.Number, sheet["D3"].Value.Kind);
        Assert.Equal(4.5, sheet["D3"].Value.AsNumber());
        Assert.True(sheet["F3"].Value.AsBoolean());
        Assert.Equal("General", sheet["D3"].NumberFormat);
    }

    [Fact]
    public void Dates_look_as_typed_dates_do()
    {
        var sheet = NewSheet();
        sheet.LoadRows("B2", TwoSales(), SaleColumns);

        Assert.Equal("3/14/2026", sheet["E3"].Text);
        Assert.Equal("m/d/yyyy", sheet["E3"].NumberFormat);
        Assert.Equal("m/d/yyyy h:mm", sheet["E4"].NumberFormat);
    }

    [Fact]
    public void A_date_takes_the_format_of_the_workbook_culture()
    {
        var sheet = NewSheet("de-DE");
        sheet.LoadRows("A1", TwoSales(), SaleColumns);
        sheet["A5"].Input = "14.03.2026";

        Assert.Equal(sheet["A5"].NumberFormat, sheet["D2"].NumberFormat);
        Assert.Equal("dd\\.mm\\.yyyy", sheet["D2"].NumberFormat);
    }

    [Fact]
    public void Null_and_DBNull_are_empty_and_clear_what_was_there()
    {
        var sheet = NewSheet();
        sheet["A2"].Value = 1;
        sheet["B2"].Value = 2;
        sheet.LoadRows("A1", new object?[] { 0 }, ("X", _ => null), ("Y", _ => DBNull.Value));

        Assert.Equal(CellValueKind.Empty, sheet["A2"].Value.Kind);
        Assert.Equal(CellValueKind.Empty, sheet["B2"].Value.Kind);
    }

    [Fact]
    public void An_empty_collection_writes_the_header_only()
    {
        var sheet = NewSheet();
        var range = sheet.LoadRows("B2", Array.Empty<Sale>(), SaleColumns);

        Assert.Equal("B2:F2", range.Address);
        Assert.Equal("Units", sheet["C2"].Text);
        Assert.Equal(CellValueKind.Empty, sheet["B3"].Value.Kind);
    }

    [Fact]
    public void Text_is_never_a_formula()
    {
        var sheet = NewSheet();
        sheet.LoadRows("A1", new[] { "=1+1" }, ("F", s => s));

        Assert.Null(sheet["A2"].Formula);
        Assert.Equal("=1+1", sheet["A2"].Value.AsText());
    }

    private enum Color { Red, Green }

    [Fact]
    public void Other_types_are_converted()
    {
        var sheet = NewSheet();
        sheet.LoadRows("A1", new[] { 0 },
            ("Enum", _ => Color.Green),
            ("Span", _ => TimeSpan.FromHours(30)),
            ("Date", _ => new DateOnly(2026, 3, 14)),
            ("Offset", _ => new DateTimeOffset(2026, 3, 14, 12, 0, 0, TimeSpan.FromHours(5))),
            ("Time", _ => new TimeOnly(12, 0)),
            ("Char", _ => 'x'),
            ("Guid", _ => 1.5f));

        Assert.Equal("Green", sheet["A2"].Value.AsText());
        Assert.Equal(1.25, sheet["B2"].Value.AsNumber());
        Assert.Equal("[h]:mm:ss", sheet["B2"].NumberFormat);
        Assert.Equal("3/14/2026", sheet["C2"].Text);
        Assert.Equal("m/d/yyyy h:mm", sheet["D2"].NumberFormat);
        Assert.Equal(0.5, sheet["E2"].Value.AsNumber());
        Assert.Equal("x", sheet["F2"].Value.AsText());
        Assert.Equal(1.5, sheet["G2"].Value.AsNumber());
    }

    [Fact]
    public void A_formula_over_the_loaded_range_calculates()
    {
        var sheet = NewSheet();
        sheet.LoadRows("B2", TwoSales(), SaleColumns);
        sheet["H1"].Formula = "=SUM(D3:D4)";
        sheet.Workbook.Recalculate();

        Assert.Equal(10.5, sheet["H1"].Value.AsNumber());
    }

    [Fact]
    public void Arguments_are_checked()
    {
        var sheet = NewSheet();
        Assert.Throws<ArgumentException>(() => sheet.LoadRows("A1", new[] { 1 }));
        Assert.Throws<ArgumentNullException>(() => sheet.LoadRows<int>("A1", [1], (null!, i => i)));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.LoadRows<int>("XFD1", [1], ("A", i => i), ("B", i => i)));
    }

    [Fact]
    public void A_value_that_cannot_be_stored_writes_nothing()
    {
        var sheet = NewSheet();
        var sales = new[]
        {
            new Sale("A", 1, 1m, new DateTime(2026, 1, 1), true),
            new Sale("B", 2, 2m, new DateTime(2026, 1, 2), true),
            new Sale("C", 3, 3m, new DateTime(1850, 1, 1), true),
        };

        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.LoadRows("B2", sales, SaleColumns));
        Assert.Null(sheet.UsedRange);

        Assert.Throws<ArgumentException>(() => sheet.LoadRows("A1", new[] { 1.0, double.NaN }, ("N", n => n)));
        Assert.Throws<ArgumentException>(() => sheet.LoadRows("A1", new[] { double.PositiveInfinity }, ("N", n => n)));
        Assert.Null(sheet.UsedRange);
    }

    [Fact]
    public void A_collection_running_past_the_last_row_writes_nothing()
    {
        var sheet = NewSheet();
        var items = Enumerable.Range(0, 1_048_576);

        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.LoadRows("A2", items, ("N", n => n)));
        Assert.Null(sheet.UsedRange);
    }
}
