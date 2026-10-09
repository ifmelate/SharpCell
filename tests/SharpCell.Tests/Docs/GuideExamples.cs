using System.Globalization;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Docs;

/// <summary>
/// The code shown in docs/guides. Everything between "// snippet:" and "// end-snippet" is copied
/// into the Markdown by tools/SharpCell.Docs; the assertions after it keep the text honest.
/// </summary>
public class GuideExamples
{
    [Fact]
    public void Build_a_workbook_in_code()
    {
        // snippet: build-in-code
        var workbook = new Workbook();
        Worksheet sheet = workbook.AddSheet("Sheet1");
        sheet["A1"].Value = 2;
        sheet["A2"].Formula = "=A1*3";
        workbook.Recalculate();
        CellValue result = sheet["A2"].Value;   // 6
        // end-snippet

        Assert.Equal(6, result.AsNumber());
    }

    // snippet: read-values
    static string Describe(CellValue value) => value.Kind switch
    {
        CellValueKind.Number => value.AsNumber().ToString(CultureInfo.InvariantCulture),
        CellValueKind.Text => value.AsText(),
        CellValueKind.Boolean => value.AsBoolean() ? "TRUE" : "FALSE",
        CellValueKind.Error => value.AsError().ToText(),   // "#DIV/0!", "#N/A", ...
        CellValueKind.Array => $"array of {value.AsArray().GetLength(0)} x {value.AsArray().GetLength(1)}",
        CellValueKind.Empty => "(empty)",
        _ => value.ToString(),
    };
    // end-snippet

    [Fact]
    public void Describe_covers_every_common_kind()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");
        Assert.Equal("6", Describe(workbook.Evaluate("=SUM(1,2,3)")));
        Assert.Equal("ab", Describe(workbook.Evaluate("=\"a\"&\"b\"")));
        Assert.Equal("TRUE", Describe(workbook.Evaluate("=1<2")));
        Assert.Equal("#DIV/0!", Describe(workbook.Evaluate("=1/0")));
        Assert.Equal("array of 2 x 3", Describe(workbook.Evaluate("=SEQUENCE(2,3)")));
        Assert.Equal("(empty)", Describe(sheet["Z9"].Value));
    }

    [Fact]
    public void Spill_and_spill_references()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");

        // snippet: spill
        sheet["C1"].Formula = "=SEQUENCE(3)";   // spills into C1:C3
        sheet["D1"].Formula = "=SUM(C1#)";      // C1# is the whole spilled range
        workbook.Recalculate();
        // sheet["C3"].Value is 3 and sheet["D1"].Value is 6
        // end-snippet

        Assert.Equal(3, sheet["C3"].Value.AsNumber());
        Assert.Null(sheet["C3"].Formula);
        Assert.Equal(6, sheet["D1"].Value.AsNumber());

        // snippet: spill-blocked
        sheet["C2"].Value = "in the way";
        workbook.Recalculate();
        // sheet["C1"].Value is #SPILL!; clear C2 and recalculate to spill again
        // end-snippet

        Assert.Equal(CellValue.Error(ErrorKind.Spill), sheet["C1"].Value);
    }

    [Fact]
    public void Names_and_lambdas()
    {
        var workbook = new Workbook();

        // snippet: names-lambda
        workbook.DefineName("TaxRate", "=0.2");
        workbook.DefineName("WithTax", "=LAMBDA(amount, amount * (1 + TaxRate))");
        CellValue gross = workbook.Evaluate("=WithTax(100)");   // 120
        // end-snippet

        Assert.Equal(120, gross.AsNumber(), 9);
    }

    [Fact]
    public void Evaluate_without_a_cell()
    {
        var workbook = new Workbook();

        // snippet: evaluate
        CellValue sum = workbook.Evaluate("=SUM(1, 2, 3)");   // 6, no sheet needed
        // end-snippet

        Assert.Equal(6, sum.AsNumber());
    }

    [Fact]
    public void A_formula_that_cannot_be_parsed()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");

        // snippet: parse-error
        string? problem = null;
        try
        {
            sheet["A1"].Formula = "=SUM(1,";
        }
        catch (FormulaParseException e)
        {
            problem = $"at position {e.Position}: {e.Message}";
        }
        // end-snippet

        Assert.NotNull(problem);
        Assert.StartsWith("at position ", problem);
    }

    [Fact]
    public void Values_are_stale_until_recalculate()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");
        sheet["A1"].Value = 2;
        sheet["A2"].Formula = "=A1*3";
        workbook.Recalculate();

        // snippet: recalculate-stale
        sheet["A1"].Value = 5;
        // sheet["A2"].Value is still 6: formulas are updated by Recalculate, not on assignment
        workbook.Recalculate();
        // sheet["A2"].Value is now 15
        // end-snippet

        Assert.Equal(15, sheet["A2"].Value.AsNumber());
    }

    [Fact]
    public void Recalculation_can_be_cancelled()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");
        sheet["A1"].Formula = "=1+1";
        var cancelled = false;

        // snippet: cancel
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        try
        {
            workbook.Recalculate(timeout.Token);
        }
        catch (OperationCanceledException)
        {
            // The workbook stays usable: the next Recalculate finishes the work.
            cancelled = true;
        }
        // end-snippet

        Assert.False(cancelled);
        Assert.Equal(2, sheet["A1"].Value.AsNumber());
    }

    [Fact]
    public void Diagnostics_report_circular_references()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");
        var lines = new List<string>();

        // snippet: diagnostics
        sheet["F1"].Formula = "=G1";
        sheet["G1"].Formula = "=F1";
        workbook.Recalculate();
        foreach (CalculationDiagnostic diagnostic in workbook.Diagnostics)
        {
            // CircularReference at Sheet1!F1: Circular reference: Sheet1!F1 -> Sheet1!G1 -> Sheet1!F1
            lines.Add($"{diagnostic.Kind} at {diagnostic.Sheet?.Name}!{diagnostic.Address}: {diagnostic.Message}");
        }
        // end-snippet

        Assert.Contains(lines, l => l.StartsWith("CircularReference at Sheet1!", StringComparison.Ordinal));
        Assert.Equal(0, sheet["F1"].Value.AsNumber());
    }

    [Fact]
    public void Culture_controls_text_to_number_conversion()
    {
        // snippet: culture
        var german = new Workbook { Culture = CultureInfo.GetCultureInfo("de-DE") };
        CellValue number = german.Evaluate("=VALUE(\"1,5\")");   // 1.5
        // end-snippet

        Assert.Equal(1.5, number.AsNumber());
    }

    [Fact]
    public void Load_from_a_stream()
    {
        var path = SampleTests.BudgetPath;

        // snippet: load-stream
        using Stream stream = File.OpenRead(path);
        Workbook workbook = XlsxReader.Load(stream);   // the stream is left open
        // end-snippet

        Assert.True(workbook.TryGetSheet("Budget", out _));
    }

    [Fact]
    public void One_workbook_per_thread()
    {
        string[] paths = [SampleTests.BudgetPath, SampleTests.BudgetPath];
        var totals = new System.Collections.Concurrent.ConcurrentBag<double>();

        // snippet: threads
        // A Workbook is not thread-safe. Separate workbooks are independent and can run in parallel.
        Parallel.ForEach(paths, path =>
        {
            Workbook workbook = XlsxReader.Load(path);
            workbook.Recalculate();
            totals.Add(workbook["Budget"]["B6"].Value.AsNumber());
        });
        // end-snippet

        Assert.Equal(new[] { 1810.0, 1810.0 }, totals.Order());
    }

    [Fact]
    public void Tables_and_structured_references()
    {
        // snippet: tables-add
        var workbook = new Workbook();
        Worksheet sheet = workbook.AddSheet("Sales");
        sheet["A1"].Value = "Region";
        sheet["B1"].Value = "Units";
        sheet["A2"].Value = "North";
        sheet["B2"].Value = 10;
        sheet["A3"].Value = "South";
        sheet["B3"].Value = 20;
        sheet.AddTable("Sales", "A1:B3");

        sheet["D1"].Formula = "=SUM(Sales[Units])";
        sheet["C2"].Formula = "=Sales[@Units]*2";   // the table's row on the formula's own row
        workbook.Recalculate();
        // end-snippet

        Assert.Equal(CellValue.Number(30), sheet["D1"].Value);
        Assert.Equal(CellValue.Number(20), sheet["C2"].Value);
    }

    [Fact]
    public void Hidden_rows()
    {
        var workbook = new Workbook();
        Worksheet sheet = workbook.AddSheet("Sheet1");
        sheet["A1"].Value = 1;
        sheet["A2"].Value = 2;
        sheet["A3"].Value = 4;

        // snippet: tables-hidden-rows
        sheet.SetRowHidden(2, true);
        sheet["B1"].Formula = "=SUBTOTAL(9,A1:A3)";     // counts a row hidden by hand
        sheet["B2"].Formula = "=SUBTOTAL(109,A1:A3)";   // skips it
        workbook.Recalculate();                         // B1 is 7, B2 is 5

        sheet.FilterMode = true;                        // as if a filter hid the row
        workbook.Recalculate();                         // both are 5
        // end-snippet

        Assert.Equal(CellValue.Number(5), sheet["B1"].Value);
        Assert.Equal(CellValue.Number(5), sheet["B2"].Value);
    }
}
