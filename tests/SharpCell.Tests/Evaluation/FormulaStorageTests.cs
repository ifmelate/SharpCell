using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Evaluation;

/// <summary>Whether a formula can be saved without the dynamic array flag, as Excel saves ordinary formulas.</summary>
public class FormulaStorageTests
{
    private static readonly string[] Plain =
    [
        "1+2", "A1*2", "SUM(A1:A9)", "VLOOKUP(A1,B:C,2,0)", "IF(A1>1,\"a\",\"b\")", "Sales[@Units]*2",
        "TaxRate*A1", "Cell*2", "SUMPRODUCT((A1:A3>1)*B1:B3)", "SUMIF(A1:A9,\">1\",B1:B9)", "ROW()-1", "ROW(A5)",
        "IFERROR(A1/B1,0)", "COUNTA(A:A)", "AVERAGE(Sales[Units])", "-A1%", "\"a\"&B2", "MAX(A1:A3)-MIN(A1:A3)",
    ];

    private static readonly string[] Dynamic =
    [
        "A1:A3*2", "SEQUENCE(3)", "ABS(A1:A3)", "SUM(A1:A3*2)", "A1#", "@A1:A3", "INDEX(A1:B5,0,1)",
        "LAMBDA(x,x)(1)", "TWICE(1)", "Range1*2", "{1,2}+1", "A1:A3", "Sales[Units]", "ROW(A1:A3)",
        "IF(A1>1,A1:A3,0)", "VLOOKUP(A1:A2,B:C,2,0)", "LET(x,1,x)", "OFFSET(A1,1,1)", "NOSUCHFUNCTION(1)",
    ];

    private static Workbook Book()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        workbook.DefineName("TaxRate", "=0.2");
        workbook.DefineName("Cell", "=S!$B$1");
        workbook.DefineName("Range1", "=S!$A$1:$A$3");
        workbook.Functions.Add("TWICE", args => args[0].AsNumber() * 2);
        sheet["J1"].Value = "Units";
        sheet["J2"].Value = 1;
        sheet.AddTable("Sales", "J1:J3");
        return workbook;
    }

    private static bool IsPlain(Workbook workbook, string formula) =>
        FormulaStorage.IsPlain(FormulaParser.Parse("=" + formula, new CellAddress(2, 10)), workbook, workbook["S"]);

    [Fact]
    public void Formulas_implicit_intersection_cannot_change_are_plain()
    {
        var workbook = Book();
        Assert.All(Plain, formula => Assert.True(IsPlain(workbook, formula), formula));
    }

    [Fact]
    public void Formulas_that_may_give_arrays_or_ranges_where_one_value_is_expected_are_dynamic()
    {
        var workbook = Book();
        Assert.All(Dynamic, formula => Assert.False(IsPlain(workbook, formula), formula));
    }

    [Fact]
    public void A_plain_formula_gives_the_same_result_read_either_way()
    {
        // Every plain formula, calculated as a legacy formula and as a dynamic one in the same place.
        var workbook = Book();
        var sheet = workbook["S"];
        for (var row = 1; row <= 9; row++)
        {
            sheet[row, 1].Value = row * 3 % 7;
            sheet[row, 2].Value = row;
            sheet[row, 3].Value = row * 10;
        }

        var column = 12;
        foreach (var formula in Plain)
        {
            var dynamicCell = sheet[2, column];
            var legacyCell = sheet[2, column + 1];
            dynamicCell.Formula = "=" + formula;
            legacyCell.SetFormula("=" + formula, legacy: true);
            column += 2;
        }

        workbook.Recalculate();
        for (var i = 0; i < Plain.Length; i++)
            Assert.Equal((Plain[i], sheet[2, 12 + 2 * i].Value), (Plain[i], sheet[2, 13 + 2 * i].Value));
    }
}
