namespace SharpCell.Tests.Evaluation;

public class TableReferenceTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    // Sales A1:D5: header, three data rows, totals row. Amount is a calculated column.
    //   Region | Units | Price | Amount
    //   North  |   10  |   2   | =[@Units]*[@Price]   20
    //   South  |   20  |   3   |                       60
    //   East   |   30  |   4   |                      120
    //   Total  | =SUBTOTAL(109,Sales[Units]) 60 |     | =SUBTOTAL(109,Sales[Amount]) 200
    public TableReferenceTests()
    {
        _s = _wb.AddSheet("S");
        string[] header = ["Region", "Units", "Price", "Amount"];
        for (var c = 0; c < header.Length; c++)
            _s[1, c + 1].Value = header[c];
        (string Region, double Units, double Price)[] rows = [("North", 10, 2), ("South", 20, 3), ("East", 30, 4)];
        for (var r = 0; r < rows.Length; r++)
        {
            _s[r + 2, 1].Value = rows[r].Region;
            _s[r + 2, 2].Value = rows[r].Units;
            _s[r + 2, 3].Value = rows[r].Price;
            _s[r + 2, 4].Formula = "=[@Units]*[@Price]";
        }

        _s["A5"].Value = "Total";
        _s["B5"].Formula = "=SUBTOTAL(109,Sales[Units])";
        _s["D5"].Formula = "=SUBTOTAL(109,Sales[Amount])";
        _s.AddTable("Sales", "A1:D5", hasTotalsRow: true);
    }

    private CellValue Calc(string formula, string cell = "F2")
    {
        _s[cell].Formula = formula;
        _wb.Recalculate();
        return _s[cell].Value;
    }

    private static CellValue N(double value) => CellValue.Number(value);

    [Theory]
    [InlineData("=SUM(Sales[Units])", 60)]
    [InlineData("=SUM(sales[UNITS])", 60)]
    [InlineData("=SUM(Sales[Amount])", 200)]
    [InlineData("=SUM(Sales[[#All],[Units]])", 120)]
    [InlineData("=COUNTA(Sales[#Headers])", 4)]
    [InlineData("=SUM(Sales[#Totals])", 260)]
    [InlineData("=SUM(Sales[[#Headers],[#Data],[Units]])", 60)]
    [InlineData("=SUM(Sales[[#Data],[#Totals],[Units]])", 120)]
    [InlineData("=SUM(Sales[[Units]:[Price]])", 69)]
    [InlineData("=SUM(Sales[[Price]:[Units]])", 69)]
    [InlineData("=SUM(Sales)", 269)]
    [InlineData("=SUM(Sales[])", 269)]
    [InlineData("=ROWS(Sales[#All])", 5)]
    [InlineData("=COLUMNS(Sales[#Data])", 4)]
    public void Areas_of_the_table(string formula, double expected)
    {
        Assert.Equal(N(expected), Calc(formula));
    }

    [Fact]
    public void Calculated_column_reads_its_own_row()
    {
        _wb.Recalculate();
        Assert.Equal([N(20), N(60), N(120), N(200)], [_s["D2"].Value, _s["D3"].Value, _s["D4"].Value, _s["D5"].Value]);
    }

    [Fact]
    public void This_row_works_beside_the_table()
    {
        Assert.Equal(N(60), Calc("=Sales[[#This Row],[Units]]*Sales[@Price]", "F3"));
        Assert.Equal(N(83), Calc("=SUM(Sales[#This Row])", "F3"));
    }

    [Theory]
    [InlineData("F1")]
    [InlineData("F5")]
    [InlineData("F6")]
    public void This_row_outside_the_data_rows_is_VALUE(string cell)
    {
        Assert.Equal(CellValue.Error(ErrorKind.Value), Calc("=Sales[@Units]", cell));
    }

    [Theory]
    [InlineData("=Sales[Nope]")]
    [InlineData("=SUM(Nope[Units])")]
    [InlineData("=[Units]")]
    public void Missing_table_or_column_is_REF(string formula)
    {
        Assert.Equal(CellValue.Error(ErrorKind.Ref), Calc(formula));
        Assert.Empty(_wb.Diagnostics);
    }

    [Fact]
    public void Table_without_header_or_totals_row()
    {
        _s["H1"].Value = 1;
        _s["H2"].Value = 2;
        _s["H3"].Value = 3;
        _s.AddTable("Bare", "H1:H3", hasHeaderRow: false);

        Assert.Equal(N(6), Calc("=SUM(Bare)"));
        Assert.Equal(N(6), Calc("=SUM(Bare[Column1])"));
        Assert.Equal(CellValue.Error(ErrorKind.Ref), Calc("=SUM(Bare[#Headers])"));
        Assert.Equal(CellValue.Error(ErrorKind.Ref), Calc("=SUM(Bare[#Totals])"));
    }

    [Fact]
    public void Changing_a_data_cell_recalculates_readers()
    {
        Assert.Equal(N(60), Calc("=SUM(Sales[Units])"));
        _s["B2"].Value = 15;
        _wb.Recalculate();
        Assert.Equal(N(65), _s["F2"].Value);
        Assert.Equal(N(30), _s["D2"].Value);
    }

    [Fact]
    public void Adding_a_table_resolves_formulas_that_named_it()
    {
        Assert.Equal(CellValue.Error(ErrorKind.Ref), Calc("=SUM(Later[X])", "J5"));
        Assert.Equal(CellValue.Error(ErrorKind.Name), Calc("=SUM(Later)", "J6"));
        _s["J1"].Value = "X";
        _s["J2"].Value = 7;
        _s.AddTable("Later", "J1:J2");
        _wb.Recalculate();
        Assert.Equal(N(7), _s["J5"].Value);
        Assert.Equal(N(7), _s["J6"].Value);
    }

    [Fact]
    public void Adding_a_table_resolves_unqualified_references_inside_it()
    {
        _s["L1"].Value = "N";
        _s["M1"].Value = "Twice";
        _s["L2"].Value = 5;
        _s["M2"].Formula = "=[@N]*2";
        _wb.Recalculate();
        Assert.Equal(CellValue.Error(ErrorKind.Ref), _s["M2"].Value);

        _s.AddTable("Pairs", "L1:M2");
        _wb.Recalculate();
        Assert.Equal(N(10), _s["M2"].Value);
    }

    [Fact]
    public void Escaped_column_name_typed_in_code()
    {
        _s["O1"].Value = "[est] Q1";
        _s["O2"].Value = 4;
        _s.AddTable("Est", "O1:O2");
        Assert.Equal(N(4), Calc("=SUM(Est['[est'] Q1])"));
    }

    [Fact]
    public void Evaluate_has_no_row_of_its_own()
    {
        Assert.Equal(N(60), _wb.Evaluate("=SUM(Sales[Units])"));
        Assert.Equal(CellValue.Error(ErrorKind.Value), _wb.Evaluate("=Sales[@Units]"));
    }

    [Fact]
    public void Evaluate_is_in_no_table_even_when_A1_is()
    {
        // A table without a header row has data in A1, where Evaluate otherwise places the formula.
        var wb = new Workbook();
        var s = wb.AddSheet("S");
        s["A1"].Value = 1;
        s["A2"].Value = 2;
        s.AddTable("Bare", "A1:A2", hasHeaderRow: false);

        Assert.Equal(N(3), wb.Evaluate("=SUM(Bare[Column1])"));
        Assert.Equal(CellValue.Error(ErrorKind.Value), wb.Evaluate("=Bare[@Column1]"));
        Assert.Equal(CellValue.Error(ErrorKind.Ref), wb.Evaluate("=SUM([Column1])"));
        Assert.Equal(CellValue.Error(ErrorKind.Value), wb.Evaluate("=Bare[#This Row]"));
    }

    [Fact]
    public void Array_result_inside_a_table_is_SPILL()
    {
        // Q3 is empty, so only the table stops the spill; outside a table the same formula spills.
        _s["Q1"].Value = "Seq";
        _s["Q2"].Formula = "=SEQUENCE(2)";
        _s["Q5"].Formula = "=SEQUENCE(1)";
        _s["S2"].Formula = "=SEQUENCE(2)";
        _s.AddTable("Spills", "Q1:Q5");
        _wb.Recalculate();
        Assert.Equal(CellValue.Error(ErrorKind.Spill), _s["Q2"].Value);
        Assert.Equal(CellValue.Empty, _s["Q3"].Value);
        Assert.Equal(N(1), _s["Q5"].Value);
        Assert.Equal(N(2), _s["S3"].Value);
    }
}
