using System.Globalization;
using SharpCell;

namespace SharpCell.Tests.Values;

public class ComparerTests
{
    [Fact]
    public void Close_numbers_are_equal()
    {
        Assert.True(NumberComparer.AreEqual(0.1 + 0.2, 0.3));
        Assert.Equal(0, NumberComparer.Compare(0.1 + 0.2, 0.3));
    }

    [Fact]
    public void Distinct_numbers_are_ordered()
    {
        Assert.False(NumberComparer.AreEqual(1, 1 + 1e-10));
        Assert.False(NumberComparer.AreEqual(0, 1e-300));
        Assert.True(NumberComparer.Compare(1, 2) < 0);
        Assert.True(NumberComparer.Compare(-1, -2) > 0);
    }

    [Fact]
    public void Text_comparison_ignores_case_but_not_accents()
    {
        var culture = CultureInfo.GetCultureInfo("en-US");
        Assert.True(TextComparer.AreEqual("abc", "ABC", culture));
        Assert.False(TextComparer.AreEqual("é", "e", culture));
        Assert.True(TextComparer.Compare("a", "B", culture) < 0);
        Assert.True(TextComparer.Compare("b", "A", culture) > 0);
    }
}
