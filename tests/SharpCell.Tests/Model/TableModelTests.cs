using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Tests.Model;

public class TableModelTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public TableModelTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = "Region";
        _s["B1"].Value = "Units";
        _s["C1"].Value = 2020;
        _s["E1"].Value = "Units";
    }

    [Fact]
    public void Columns_come_from_the_header_row()
    {
        var table = _s.AddTable("Sales", "A1:E4");

        Assert.Equal(["Region", "Units", "2020", "Column4", "Units2"], table.Columns);
        Assert.Equal("Sales", table.Name);
        Assert.Same(_s, table.Worksheet);
        Assert.Equal("A1:E4", table.Range);
        Assert.True(table.HasHeaderRow);
        Assert.False(table.HasTotalsRow);
    }

    [Fact]
    public void Table_without_header_row_has_numbered_columns()
    {
        var table = _s.AddTable("Bare", "G1:H2", hasHeaderRow: false, hasTotalsRow: true);
        Assert.Equal(["Column1", "Column2"], table.Columns);
        Assert.True(table.HasTotalsRow);
    }

    [Fact]
    public void Columns_cannot_be_changed_from_outside()
    {
        var table = _s.AddTable("Sales", "A1:B3");
        Assert.Throws<NotSupportedException>(() => ((IList<string>)table.Columns)[0] = "Changed");

        var names = new List<string> { "X" };
        var other = _s.AddTable("Other", new Area(10, 1, 11, 1), hasHeaderRow: true, hasTotalsRow: false, names);
        names[0] = "Changed";
        Assert.Equal(["X"], other.Columns);
    }

    [Fact]
    public void Tables_are_found_by_name_ignoring_case()
    {
        var table = _s.AddTable("Sales", "A1:B3");
        Assert.True(_wb.TryGetTable("SALES", out var found));
        Assert.Same(table, found);
        Assert.False(_wb.TryGetTable("Other", out _));
        Assert.Equal([table], _wb.Tables);
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("1x")]
    [InlineData("TRUE")]
    [InlineData("")]
    public void Invalid_table_name_is_rejected(string name)
    {
        Assert.Throws<ArgumentException>(() => _s.AddTable(name, "A1:B3"));
    }

    [Fact]
    public void Table_and_defined_names_share_one_namespace()
    {
        _wb.DefineName("Rate", "0.2");
        Assert.Throws<ArgumentException>(() => _s.AddTable("rate", "A1:B3"));

        _s.AddTable("Sales", "A1:B3");
        Assert.Throws<ArgumentException>(() => _wb.DefineName("SALES", "1"));
        Assert.Throws<ArgumentException>(() => _wb.DefineName("Sales", "1", _s));
        Assert.Throws<ArgumentException>(() => _s.AddTable("sales", "J1:K3"));
    }

    [Fact]
    public void Tables_cannot_overlap()
    {
        _s.AddTable("Sales", "A1:B3");
        Assert.Throws<ArgumentException>(() => _s.AddTable("Other", "B3:C5"));
        _wb.AddSheet("T").AddTable("Elsewhere", "A1:B3");
    }

    [Theory]
    [InlineData("A1:B1", true, false)]
    [InlineData("A1:B2", true, true)]
    [InlineData("A1:B1", false, true)]
    public void Table_needs_a_data_row(string range, bool header, bool totals)
    {
        Assert.Throws<ArgumentException>(() => _s.AddTable("Sales", range, header, totals));
    }

    [Fact]
    public void Single_cell_table_without_header_row()
    {
        var table = _s.AddTable("One", "G5", hasHeaderRow: false);
        Assert.Equal("G5:G5", table.Range);
        Assert.Equal(["Column1"], table.Columns);
    }

    [Theory]
    [InlineData("A:A")]
    [InlineData("1:1")]
    [InlineData("nonsense")]
    [InlineData("Other!A1:B2")]
    public void Range_must_be_an_area_on_the_sheet(string range)
    {
        Assert.Throws<ArgumentException>(() => _s.AddTable("Sales", range));
    }
}
