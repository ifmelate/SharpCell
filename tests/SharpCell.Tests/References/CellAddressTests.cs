using SharpCell;

namespace SharpCell.Tests.References;

public class CellAddressTests
{
    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(16384, "XFD")]
    public void Column_names_round_trip(int column, string name)
    {
        Assert.Equal(name, ColumnNames.ToName(column));
        Assert.True(ColumnNames.TryParse(name, out var parsed));
        Assert.Equal(column, parsed);
        Assert.True(ColumnNames.TryParse(name.ToLowerInvariant(), out parsed));
        Assert.Equal(column, parsed);
    }

    [Theory]
    [InlineData("")]
    [InlineData("XFE")]
    [InlineData("AAAA")]
    [InlineData("A1")]
    public void Invalid_column_names_are_rejected(string name)
    {
        Assert.False(ColumnNames.TryParse(name, out _));
    }

    [Fact]
    public void Cell_address_formats_and_parses()
    {
        var address = new CellAddress(3, 2);
        Assert.Equal("B3", address.ToString());
        Assert.True(CellAddress.TryParse("b3", out var parsed));
        Assert.Equal(address, parsed);
    }

    [Theory]
    [InlineData("A0")]
    [InlineData("A1048577")]
    [InlineData("XFE1")]
    [InlineData("$A$1")]
    [InlineData("1A")]
    [InlineData("A")]
    public void Invalid_cell_addresses_are_rejected(string text)
    {
        Assert.False(CellAddress.TryParse(text, out _));
    }

    [Fact]
    public void Relative_axis_wraps_around_the_sheet()
    {
        Assert.Equal(CellAddress.MaxRow, AxisRef.Relative(-1).Resolve(1, CellAddress.MaxRow));
        Assert.Equal(1, AxisRef.Relative(1).Resolve(CellAddress.MaxRow, CellAddress.MaxRow));
        Assert.Equal(5, AxisRef.Absolute(5).Resolve(100, CellAddress.MaxRow));
    }
}
