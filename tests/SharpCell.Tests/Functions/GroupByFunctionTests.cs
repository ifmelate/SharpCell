namespace SharpCell.Tests.Functions;

public class GroupByFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _d;

    public GroupByFunctionTests()
    {
        _d = _wb.AddSheet("Data");
        string[] region = ["North", "South", "North", "South", "North", "South"];
        string[] product = ["A", "B", "A", "C", "B", "A"];
        double[] units = [10, 20, 30, 40, 50, 60];
        _d["A1"].Value = "Region";
        _d["B1"].Value = "Product";
        _d["C1"].Value = "Units";
        for (var i = 0; i < 6; i++)
        {
            _d[i + 2, 1].Value = region[i];
            _d[i + 2, 2].Value = product[i];
            _d[i + 2, 3].Value = units[i];
        }
    }

    [Fact]
    public void PERCENTOF_is_the_sum_of_a_part_over_the_sum_of_all()
    {
        Assert.Equal(60.0 / 210, _wb.Evaluate("=PERCENTOF(Data!C2:C4,Data!C2:C7)").AsNumber(), 12);
        Assert.Equal(0.5, _wb.Evaluate("=PERCENTOF(5,10)").AsNumber(), 12);
    }

    [Theory]
    [InlineData("=PERCENTOF({1,2},{0,0})", "#DIV/0!")]
    [InlineData("=PERCENTOF({1,#N/A},{4})", "#NUM!")]   // Excel for the web: an error inside gives #NUM!
    [InlineData("=PERCENTOF({1},{#NUM!})", "#NUM!")]
    [InlineData("=PERCENTOF(\"x\",4)", "#VALUE!")]
    public void PERCENTOF_errors(string formula, string expected)
    {
        Assert.Equal(expected, _wb.Evaluate(formula).ToString());
    }

    [Fact]
    public void PERCENTOF_ignores_text_in_ranges_and_arrays_as_SUM_does()
    {
        Assert.Equal(0.25, _wb.Evaluate("=PERCENTOF({1,\"x\"},{4})").AsNumber(), 12);
        Assert.Equal(0.0, _wb.Evaluate("=PERCENTOF(Data!B2:B7,Data!C2:C7)").AsNumber(), 12);
    }

    // The data of tests/corpus/excel-web/groupby.xlsx; the expected grids are Excel's.
    private static Workbook Sales()
    {
        var workbook = new Workbook();
        var data = workbook.AddSheet("D");
        string[] regions = ["North", "South", "North", "South", "North", "South"];
        string?[] products = ["A", "B", "A", "C", null, "B"];
        double[] units = [10, 20, 30, 40, 50, 60];
        for (var i = 0; i < 6; i++)
        {
            data[i + 1, 1].Value = regions[i];
            if (products[i] is { } product)
                data[i + 1, 2].Value = product;
            data[i + 1, 3].Value = units[i];
        }

        return workbook;
    }

    private static string Grid(Workbook workbook, string formula)
    {
        var value = workbook.Evaluate(formula);
        if (value.Kind != CellValueKind.Array)
            return value.ToString();
        var array = value.AsArray();
        return string.Join(" / ", Enumerable.Range(0, array.GetLength(0))
            .Select(r => string.Join(",", Enumerable.Range(0, array.GetLength(1)).Select(c => array[r, c].ToString()))));
    }

    [Fact]
    public void GROUPBY_groups_sorts_and_totals()
    {
        Assert.Equal("\"North\",90 / \"South\",120 / \"Total\",210", Grid(Sales(), "=GROUPBY(D!A1:A6,D!C1:C6,SUM)"));
    }

    [Fact]
    public void GROUPBY_subtotals_say_grand_total_and_empty_keys_come_last()
    {
        Assert.Equal(
            "\"North\",\"A\",40 / \"North\",<empty>,50 / \"North\",\"\",90 / \"South\",\"B\",80 / \"South\",\"C\",40 / \"South\",\"\",120 / \"Grand Total\",\"\",210",
            Grid(Sales(), "=GROUPBY(D!A1:B6,D!C1:C6,SUM,0,2)"));
    }

    [Fact]
    public void GROUPBY_with_several_functions_names_them()
    {
        Assert.Equal("\"\",\"SUM\",\"MAX\" / \"North\",90,50 / \"South\",120,60 / \"Total\",210,60",
            Grid(Sales(), "=GROUPBY(D!A1:A6,D!C1:C6,HSTACK(SUM,MAX))"));
    }

    [Fact]
    public void PIVOTBY_leaves_empty_text_where_no_row_falls_and_totals_both_ways()
    {
        Assert.Equal(
            "\"\",\"A\",\"B\",\"C\",<empty>,\"Total\" / \"North\",40,\"\",\"\",50,90 / \"South\",\"\",80,40,\"\",120 / \"Total\",40,80,40,50,210",
            Grid(Sales(), "=PIVOTBY(D!A1:A6,D!B1:B6,D!C1:C6,SUM)"));
    }

    [Theory]
    [InlineData("=GROUPBY(D!A1:A6,D!C1:C5,SUM)")]
    [InlineData("=GROUPBY(D!A1:A6,D!C1:C6,SUM,0,1,1,D!A1:A5=\"North\")")]
    [InlineData("=GROUPBY(D!A1:A6,D!C1:C6,1)")]
    [InlineData("=GROUPBY(D!A1:A6,D!C1:C6,SUM,7)")]
    [InlineData("=GROUPBY(D!A1:A6,D!C1:C6,SUM,0,2)")]
    [InlineData("=GROUPBY(D!A1:A6,D!C1:C6,SUM,0,1,5)")]
    [InlineData("=GROUPBY(D!A1:B6,D!C1:C6,SUM,0,1,{-1,3})")]
    [InlineData("=GROUPBY(D!A1:B6,D!C1:C6,SUM,0,2,,,1)")]
    [InlineData("=GROUPBY(D!A1:A6,D!C1:C6,SUM,0,1,1,D!A1:A6=\"East\")")]
    [InlineData("=PIVOTBY(D!A1:A6,D!B1:B6,D!B1:C6,SUM)")]
    [InlineData("=PIVOTBY(D!A1:A6,D!B1:B6,D!C1:C6,SUM,0,1,1,1,1,,9)")]
    public void Mismatched_lengths_are_value_errors(string formula)
    {
        Assert.Equal(CellValue.Error(ErrorKind.Value), Sales().Evaluate(formula));
    }

    private static Workbook Many(int rows, int groups)
    {
        var workbook = new Workbook();
        var data = workbook.AddSheet("D");
        for (var r = 1; r <= rows; r++)
        {
            data[r, 1].Value = "Key" + (r % groups);
            data[r, 2].Value = r;
        }

        return workbook;
    }

    [Fact]
    public void GROUPBY_handles_thousands_of_distinct_keys_quickly()
    {
        var workbook = Many(20_000, 5_000);
        var clock = System.Diagnostics.Stopwatch.StartNew();

        var result = workbook.Evaluate("=GROUPBY(D!A1:A20000,D!B1:B20000,SUM,0,1,-2)");

        Assert.Equal(5_001, result.AsArray().GetLength(0));
        Assert.True(clock.Elapsed < TimeSpan.FromSeconds(3), $"took {clock.Elapsed}");
    }

    [Fact]
    public void GROUPBY_stops_when_the_calculation_is_cancelled()
    {
        var workbook = Many(50_000, 10);
        using var cancel = new CancellationTokenSource();
        workbook.Functions.Add("CANCELNOW", args =>
        {
            cancel.Cancel();
            return CellValue.Array(args[0].AsArray());
        });

        Assert.ThrowsAny<OperationCanceledException>(() =>
            workbook.Evaluate("=GROUPBY(CANCELNOW(D!A1:A50000),D!B1:B50000,LAMBDA(x,1))", cancel.Token));
    }

    [Fact]
    public void A_lambda_with_two_parameters_gets_the_values_its_group_is_a_part_of()
    {
        var workbook = Sales();
        Assert.Equal(Grid(workbook, "=GROUPBY(D!A1:A6,D!C1:C6,PERCENTOF,0,0)"),
            Grid(workbook, "=GROUPBY(D!A1:A6,D!C1:C6,LAMBDA(part,all,SUM(part)/SUM(all)),0,0)"));
        Assert.Equal(Grid(workbook, "=PIVOTBY(D!A1:A6,D!B1:B6,D!C1:C6,PERCENTOF,0,1,1,1,1,,1)"),
            Grid(workbook, "=PIVOTBY(D!A1:A6,D!B1:B6,D!C1:C6,LAMBDA(part,all,SUM(part)/SUM(all)),0,1,1,1,1,,1)"));
    }
}
