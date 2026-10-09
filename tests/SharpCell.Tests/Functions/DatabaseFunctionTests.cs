using System;
using SharpCell;

namespace SharpCell.Tests.Functions;

/// <summary>The database functions on the sample of Excel's documentation (criteria in A1:F3, data in A4:E10).</summary>
public class DatabaseFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public DatabaseFunctionTests()
    {
        _s = _wb.AddSheet("S");
        string[] criteriaHeader = ["Tree", "Height", "Age", "Yield", "Profit", "Height"];
        for (var c = 0; c < criteriaHeader.Length; c++)
            _s[1, c + 1].Value = criteriaHeader[c];
        _s["A2"].Value = "=Apple";
        _s["B2"].Value = ">10";
        _s["F2"].Value = "<16";
        _s["A3"].Value = "=Pear";

        object[][] data =
        [
            ["Tree", "Height", "Age", "Yield", "Profit"],
            ["Apple", 18, 20, 14, 105],
            ["Pear", 12, 12, 10, 96],
            ["Cherry", 13, 14, 9, 105],
            ["Apple", 14, 15, 10, 75],
            ["Pear", 9, 8, 8, 76.8],
            ["Apple", 8, 9, 6, 45],
        ];
        for (var r = 0; r < data.Length; r++)
        {
            for (var c = 0; c < data[r].Length; c++)
            {
                _s[r + 4, c + 1].Value = data[r][c] switch
                {
                    string text => CellValue.Text(text),
                    int number => CellValue.Number(number),
                    var number => CellValue.Number((double)number),
                };
            }
        }

        // H1:H2 plain text (begins with), I1:I3 an empty condition row, J1:K2 a computed criterion.
        _s["H1"].Value = "Tree";
        _s["H2"].Value = "app";
        _s["I1"].Value = "Tree";
        _s["I2"].Value = "Cherry";
        _s["J1"].Value = "Big";
        _s["J2"].Value = true;
        _s["K1"].Value = "Big";
        _s["K2"].Value = false;
        _s["L1"].Value = "Profit";
        _s["L2"].Value = 105;
    }

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    public static TheoryData<string, CellValue> Documented => new()
    {
        { "=DCOUNT(A4:E10,\"Age\",A1:F2)", 1 },
        { "=DCOUNTA(A4:E10,\"Profit\",A1:F2)", 1 },
        { "=DMAX(A4:E10,\"Profit\",A1:F3)", 96 },
        { "=DMIN(A4:E10,\"Profit\",A1:B2)", 75 },
        { "=DSUM(A4:E10,\"Profit\",A1:A2)", 225 },
        { "=DSUM(A4:E10,\"Profit\",A1:F2)", 75 },
        { "=DPRODUCT(A4:E10,\"Yield\",A1:A2)", 14 * 10 * 6 },
        { "=DAVERAGE(A4:E10,\"Yield\",A1:B2)", 12 },
        { "=DAVERAGE(A4:E10,3,A4:E10)", 13 },
        { "=DVAR(A4:E10,\"Yield\",A1:A3)", 8.8 },
        { "=DVARP(A4:E10,\"Yield\",A1:A3)", 7.04 },
        { "=DSTDEV(A4:E10,\"Yield\",A1:A3)", Math.Sqrt(8.8) },
        { "=DSTDEVP(A4:E10,\"Yield\",A1:A3)", Math.Sqrt(7.04) },
        { "=DGET(A4:E10,\"Yield\",A1:F2)", 10 },
        { "=DGET(A4:E10,\"Yield\",A1:A3)", Err(ErrorKind.Num) },
    };

    public static TheoryData<string, CellValue> Edges => new()
    {
        // Fields: names in any case, numbers truncated, logical values as numbers.
        { "=DSUM(A4:E10,\"PROFIT\",A1:A2)", 225 },
        { "=DSUM(A4:E10,5.9,A1:A2)", 225 },
        { "=DSUM(A4:E10,E4,A1:A2)", 225 },
        { "=DCOUNTA(A4:E10,TRUE,A1:A2)", 3 },
        { "=DSUM(A4:E10,0,A1:A2)", Err(ErrorKind.Value) },
        { "=DSUM(A4:E10,6,A1:A2)", Err(ErrorKind.Value) },
        { "=DSUM(A4:E10,\"Price\",A1:A2)", Err(ErrorKind.Value) },
        { "=DSUM(A4:E10,D4:E4,A1:A2)", Err(ErrorKind.Value) },
        { "=DSUM(A4:E10,Z99,A1:A2)", Err(ErrorKind.Value) },
        { "=DSUM(A4:E10,1/0,A1:A2)", Err(ErrorKind.Div0) },
        { "=DCOUNT(A4:E10,,A1:A2)", 3 },
        { "=DCOUNTA(A4:E10,,A1:A3)", 5 },
        { "=DSUM(A4:E10,,A1:A2)", Err(ErrorKind.Value) },

        // Criteria: plain text begins with, numbers equal, an empty row selects everything.
        { "=DCOUNT(A4:E10,\"Yield\",H1:H2)", 3 },
        { "=DSUM(A4:E10,\"Yield\",L1:L2)", 23 },
        { "=DCOUNT(A4:E10,\"Yield\",I1:I3)", 6 },
        { "=DCOUNT(A4:E10,\"Yield\",I1:I2)", 1 },

        // A computed criterion counts by its value: TRUE selects every record.
        { "=DCOUNT(A4:E10,\"Yield\",J1:J2)", 6 },
        { "=DCOUNT(A4:E10,\"Yield\",K1:K2)", 0 },

        // Shapes and empty results.
        { "=DSUM(A4:E4,\"Yield\",A1:A2)", Err(ErrorKind.Value) },
        { "=DSUM(A4:E10,\"Yield\",A1:B1)", Err(ErrorKind.Value) },
        { "=DSUM(A5:E10,\"Yield\",A1:A2)", Err(ErrorKind.Value) },
        { "=DSUM(A4:E10,\"Yield\",(A1:A2,B1:B2))", Err(ErrorKind.Value) },
        { "=DSUM(A4:E10,\"Tree\",A1:A2)", 0 },
        { "=DMAX(A4:E10,\"Tree\",A1:A2)", 0 },
        { "=DPRODUCT(A4:E10,\"Tree\",A1:A2)", 0 },
        { "=DAVERAGE(A4:E10,\"Tree\",A1:A2)", Err(ErrorKind.Div0) },
        { "=DSTDEV(A4:E10,\"Yield\",I1:I2)", Err(ErrorKind.Div0) },
        { "=DSTDEVP(A4:E10,\"Yield\",I1:I2)", 0 },
        { "=DVARP(A4:E10,\"Yield\",K1:K2)", Err(ErrorKind.Div0) },
        { "=DGET(A4:E10,\"Yield\",K1:K2)", Err(ErrorKind.Value) },
        { "=DGET(A4:E10,\"Tree\",I1:I2)", "Cherry" },
    };

    [Theory]
    [MemberData(nameof(Documented))]
    [MemberData(nameof(Edges))]
    public void Database_functions(string formula, CellValue expected)
    {
        var actual = _wb.Evaluate(formula);
        if (expected.Kind == CellValueKind.Number && actual.Kind == CellValueKind.Number)
            Assert.Equal(expected.AsNumber(), actual.AsNumber(), 12);
        else
            Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData(">=2", "Apple", false)]
    [InlineData("=a", "Apple", false)]
    [InlineData("=apple", "Apple", true)]
    [InlineData("a", "Apple", true)]
    [InlineData("a", "Banana", false)]
    [InlineData("*an", "Banana", true)]
    [InlineData("<>Apple", "Apple", false)]
    [InlineData("<>Apple", "Apples", true)]
    [InlineData("=", "", true)]
    [InlineData("=", "x", false)]
    [InlineData(">m", "Pear", true)]
    [InlineData("TRUE", "x", false)]
    public void Text_conditions(string condition, string value, bool selected)
    {
        var s = _wb.AddSheet("T");
        s["A1"].Value = "F";
        s["A2"].Value = value.Length == 0 ? CellValue.Empty : CellValue.Text(value);
        s["B2"].Value = 1;
        s["B1"].Value = "N";
        s["D1"].Value = "F";
        s["D2"].Value = condition;
        Assert.Equal(CellValue.Number(selected ? 1 : 0), _wb.Evaluate("=DSUM(T!A1:B2,\"N\",T!D1:D2)"));
    }

    [Fact]
    public void Whole_column_database_reads_only_stored_rows()
    {
        var s = _wb.AddSheet("Big");
        s["A1"].Value = "K";
        s["B1"].Value = "V";
        for (var row = 2; row <= 101; row++)
        {
            s[row, 1].Value = row % 2 == 0 ? "even" : "odd";
            s[row, 2].Value = row;
        }

        s["D1"].Value = "K";
        s["D2"].Value = "even";
        Assert.Equal(CellValue.Number(50), _wb.Evaluate("=DCOUNT(Big!A:B,\"V\",Big!D1:D2)"));
        Assert.Equal(CellValue.Number(2550), _wb.Evaluate("=DSUM(Big!A:B,2,Big!D1:D2)"));
    }

    [Fact]
    public void Formula_records_and_conditions_are_calculated_first()
    {
        _s["N4"].Value = "Total";
        _s["N5"].Formula = "=E5*2";
        _s["N6"].Formula = "=E6*2";
        _s["P1"].Value = "Total";
        _s["P2"].Formula = "=\">\"&100";
        _s["Q1"].Formula = "=DSUM(A4:N10,\"Total\",P1:P2)";
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(402), _s["Q1"].Value);

        _s["E6"].Value = 40;
        _wb.Recalculate();
        Assert.Equal(CellValue.Number(210), _s["Q1"].Value);
    }
}
