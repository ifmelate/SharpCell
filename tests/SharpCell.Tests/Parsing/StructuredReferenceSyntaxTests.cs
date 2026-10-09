using SharpCell.Parsing;

namespace SharpCell.Tests.Parsing;

public class StructuredReferenceSyntaxTests
{
    private const TableRows HeadersAndData = TableRows.Headers | TableRows.Data;
    private const TableRows DataAndTotals = TableRows.Data | TableRows.Totals;

    [Theory]
    [InlineData("Sales[]", "Sales", TableRows.Data, null, null)]
    [InlineData("Sales[Units]", "Sales", TableRows.Data, "Units", "Units")]
    [InlineData("Sales[#All]", "Sales", TableRows.All, null, null)]
    [InlineData("Sales[#Data]", "Sales", TableRows.Data, null, null)]
    [InlineData("Sales[#Headers]", "Sales", TableRows.Headers, null, null)]
    [InlineData("Sales[#Totals]", "Sales", TableRows.Totals, null, null)]
    [InlineData("Sales[#This Row]", "Sales", TableRows.ThisRow, null, null)]
    [InlineData("Sales[#this row]", "Sales", TableRows.ThisRow, null, null)]
    [InlineData("Sales[[#This Row],[Jan]:[Dec]]", "Sales", TableRows.ThisRow, "Jan", "Dec")]
    [InlineData("Sales[[#Totals], [Year]]", "Sales", TableRows.Totals, "Year", "Year")]
    [InlineData("Sales[[#Data],[Jan]:[Mar]]", "Sales", TableRows.Data, "Jan", "Mar")]
    [InlineData("Sales[[#Headers],[#Data],[Units]]", "Sales", HeadersAndData, "Units", "Units")]
    [InlineData("Sales[[#Data],[#Totals]]", "Sales", DataAndTotals, null, null)]
    [InlineData("Sales[[Units]:[Price]]", "Sales", TableRows.Data, "Units", "Price")]
    [InlineData("Sales[[the two ]]", "Sales", TableRows.Data, "the two ", "the two ")]
    [InlineData("Sales[@Units]", "Sales", TableRows.ThisRow, "Units", "Units")]
    [InlineData("Sales[@[Unit Cost]]", "Sales", TableRows.ThisRow, "Unit Cost", "Unit Cost")]
    [InlineData("Sales[@[A]:[B]]", "Sales", TableRows.ThisRow, "A", "B")]
    [InlineData("Sales[@]", "Sales", TableRows.ThisRow, null, null)]
    [InlineData("[Units]", null, TableRows.Data, "Units", "Units")]
    [InlineData("[@Units]", null, TableRows.ThisRow, "Units", "Units")]
    // Escapes from tables.xlsx (Table9): ' before [ ] # '.
    [InlineData("Table9['#Shame]", "Table9", TableRows.Data, "#Shame", "#Shame")]
    [InlineData("Table9['[Not me']]", "Table9", TableRows.Data, "[Not me]", "[Not me]")]
    [InlineData("Table9[Hi, ''This]", "Table9", TableRows.Data, "Hi, 'This", "Hi, 'This")]
    [InlineData("Table9[[#This Row],[Hi, ''This]:['[Not me']]]", "Table9", TableRows.ThisRow, "Hi, 'This", "[Not me]")]
    [InlineData("Table9[Totals '[']^&*@'#$,,<>]", "Table9", TableRows.Data, "Totals []^&*@#$,,<>", "Totals []^&*@#$,,<>")]
    // TableRows is internal, so the theory takes it as object.
    public void Parses(string text, string? table, object rows, string? first, string? last)
    {
        var parsed = StructuredReferenceSyntax.Parse(text, 0);
        Assert.Equal(new StructuredReference(table, (TableRows)rows, first, last), parsed);
    }

    [Theory]
    [InlineData("Sales[[#All],[#Data]]")]
    [InlineData("Sales[[#Headers],[#Totals]]")]
    [InlineData("Sales[[#This Row],[#Data]]")]
    [InlineData("Sales[[#Data],[#Data]]")]
    [InlineData("Sales[[A],[B]]")]
    [InlineData("Sales[[A]:[#Totals]]")]
    [InlineData("Sales[#Everything]")]
    [InlineData("Sales[[]]")]
    [InlineData("Sales[[A] [B]]")]
    [InlineData("Sales[@[#Totals]]")]
    [InlineData("Sales[A[B]]")]
    public void Rejects(string text)
    {
        Assert.Throws<FormulaParseException>(() => StructuredReferenceSyntax.Parse(text, 0));
    }

    [Fact]
    public void Parser_keeps_the_text_and_the_parsed_reference()
    {
        var node = Assert.IsType<StructuredReferenceNode>(FormulaParser.Parse("=Sales[@Units]", new CellAddress(1, 1)));
        Assert.Equal("Sales[@Units]", node.Text);
        Assert.Equal(new StructuredReference("Sales", TableRows.ThisRow, "Units", "Units"), node.Reference);
    }

    [Fact]
    public void Invalid_structured_reference_fails_the_formula()
    {
        Assert.Throws<FormulaParseException>(() => FormulaParser.Parse("=SUM(Sales[[#All],[#Data]])", new CellAddress(1, 1)));
    }

    [Fact]
    public void Formula_text_is_kept_as_written()
    {
        var wb = new Workbook();
        var s = wb.AddSheet("S");
        s["A1"].Formula = "=Sales[@Units]*2+SUM(Sales[[#Totals],[Units]:[Price]])";
        Assert.Equal("=Sales[@Units]*2+SUM(Sales[[#Totals],[Units]:[Price]])", s["A1"].Formula);
    }
}
