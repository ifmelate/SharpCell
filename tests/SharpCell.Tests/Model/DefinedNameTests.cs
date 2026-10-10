using System.Linq;

namespace SharpCell.Tests.Model;

public class DefinedNameTests
{
    [Fact]
    public void DefinedNames_lists_names_as_written_with_their_scope()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        workbook.DefineName("TaxRate", "0.2");
        workbook.DefineName("Local", "=S!A1", sheet);

        var names = workbook.DefinedNames;

        Assert.Equal(new[] { "TaxRate", "Local" }, names.Select(n => n.Name));
        Assert.Equal(new[] { "=0.2", "=S!A1" }, names.Select(n => n.Formula));
        Assert.Null(names[0].Scope);
        Assert.Same(sheet, names[1].Scope);
        Assert.Equal("TaxRate", names[0].ToString());
    }

    [Fact]
    public void Redefining_a_name_keeps_its_place_and_takes_the_new_spelling()
    {
        var workbook = new Workbook();
        workbook.DefineName("Rate", "1");
        workbook.DefineName("Other", "2");
        workbook.DefineName("RATE", "3");

        Assert.Equal(new[] { ("RATE", "=3"), ("Other", "=2") }, workbook.DefinedNames.Select(n => (n.Name, n.Formula)));
    }

    [Fact]
    public void The_same_name_in_two_scopes_is_listed_twice()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        workbook.DefineName("Rate", "1");
        workbook.DefineName("Rate", "2", sheet);

        Assert.Equal(2, workbook.DefinedNames.Count);
    }

    [Fact]
    public void DefinedNames_is_a_snapshot()
    {
        var workbook = new Workbook();
        var before = workbook.DefinedNames;
        workbook.DefineName("Rate", "1");

        Assert.Empty(before);
        Assert.Single(workbook.DefinedNames);
    }

    [Fact]
    public void A_name_from_a_file_that_could_not_be_parsed_keeps_its_text()
    {
        using var file = new Xlsx.TestXlsx().Sheet("S", "").Name("Ext", "[1]Sheet1!$A$1").Build();
        var workbook = SharpCell.Xlsx.XlsxReader.Load(file);

        var name = Assert.Single(workbook.DefinedNames);
        Assert.Equal("Ext", name.Name);
        Assert.Equal("=[1]Sheet1!$A$1", name.Formula);
    }
}
