using SharpCell.Xlsx;

namespace SharpCell.Tests.Docs;

/// <summary>
/// The code shown in docs/guides/cells.md, custom-functions.md and the parts of recalculation.md
/// about changes, precedents and clones. See <see cref="GuideExamples"/> for how snippets work.
/// </summary>
public class AppGuideExamples
{
    private static (Workbook Workbook, Worksheet Sheet) Budget()
    {
        var workbook = XlsxReader.Load(SampleTests.BudgetPath);
        workbook.Recalculate();
        return (workbook, workbook["Budget"]);
    }

    [Fact]
    public void Enumerate_cells()
    {
        var (_, sheet) = Budget();
        var lines = new List<string>();

        // snippet: cells-enumerate
        foreach (Cell cell in sheet.Cells)   // non-empty cells, row by row
        {
            if (cell.Formula is not null)
                lines.Add($"{cell.Address}: {cell.Formula} = {cell.Value}");
        }

        CellRange? used = sheet.UsedRange;   // A1:F6 for the sample budget; null for an empty sheet
        // end-snippet

        Assert.Contains("B6: =SUM(B2:B5) = 1810", lines);
        Assert.Equal("A1:F6", used!.Address);
    }

    [Fact]
    public void Read_and_write_blocks()
    {
        var (workbook, sheet) = Budget();

        // snippet: cells-values
        CellRange amounts = sheet.Range("B2:B5");
        CellValue[,] before = amounts.GetValues();   // [row, column] from B2: 1200, 450, 120, 40

        amounts.SetValues(new CellValue[,] { { 1300 }, { 500 }, { 150 }, { 50 } });
        workbook.Recalculate();
        // sheet["B6"].Value is now 2000
        // end-snippet

        Assert.Equal(1200, before[0, 0].AsNumber());
        Assert.Equal(2000, sheet["B6"].Value.AsNumber());
    }

    [Fact]
    public void Dates()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");

        // snippet: cells-dates
        sheet["A1"].Value = CellValue.DateTime(new DateTime(2026, 3, 15), workbook.DateSystem);
        sheet["A2"].Formula = "=EOMONTH(A1, 0)";
        workbook.Recalculate();

        DateTime monthEnd = sheet["A2"].Value.AsDateTime(workbook.DateSystem);   // 2026-03-31
        // end-snippet

        Assert.Equal(new DateTime(2026, 3, 31), monthEnd);
    }

    [Fact]
    public void Display_text()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");

        // snippet: cells-text
        sheet["A1"].Value = 1234.5;
        sheet["A1"].NumberFormat = "#,##0.00";   // as read from the file, or set here
        string shown = sheet["A1"].Text;          // "1,234.50"
        // end-snippet

        Assert.Equal("1,234.50", shown);
    }

    [Fact]
    public void Cell_styles()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");

        // snippet: cells-styles
        CellStyle header = sheet["A1"].Style with   // the workbook's default until the cell has its own
        {
            Font = sheet["A1"].Style.Font with { Bold = true },
            Fill = CellColor.FromRgb(0xD9E1F2),
            BottomBorder = new CellBorder(CellBorderStyle.Thin),
            HorizontalAlignment = CellHorizontalAlignment.Center,
        };
        sheet["A1"].Style = header;
        bool bold = sheet["A1"].Style.Font.Bold;      // true
        string font = sheet["B1"].Style.Font.Name;    // "Calibri", the default
        // end-snippet

        Assert.True(bold);
        Assert.Equal("Calibri", font);
        Assert.Equal("#D9E1F2", sheet["A1"].Style.Fill.ToString());
    }

    [Fact]
    public void Sheet_geometry()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");

        // snippet: cells-geometry
        sheet.SetColumnWidth(2, 20);      // column B, in the units .xlsx files use
        sheet.SetRowHeight(1, 30);        // row 1, in points
        sheet.SetColumnHidden(3, true);   // column C
        sheet.Merge("A1:D1");
        double? width = sheet.ColumnWidth(1);           // null: column A has the default width
        string merged = sheet.MergedAreas[0].Address;   // "A1:D1"
        // end-snippet

        Assert.Null(width);
        Assert.Equal("A1:D1", merged);
    }

    [Fact]
    public void Sheet_view()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");

        // snippet: cells-view
        sheet.FrozenRows = 1;          // the header row stays in place while the rest scrolls
        sheet.FrozenColumns = 1;       // and so does column A
        sheet.ShowGridlines = false;   // as View > Gridlines turned off in Excel
        // end-snippet

        Assert.Equal((1, 1, false), (sheet.FrozenRows, sheet.FrozenColumns, sheet.ShowGridlines));
    }

    [Fact]
    public void Defined_names()
    {
        var workbook = new Workbook();
        workbook.AddSheet("Sheet1");
        workbook.DefineName("TaxRate", "=0.2");
        var lines = new List<string>();

        // snippet: cells-names
        foreach (DefinedName name in workbook.DefinedNames)
            lines.Add($"{name.Name} {name.Formula} {name.Scope?.Name ?? "(workbook)"}");   // TaxRate =0.2 (workbook)
        // end-snippet

        Assert.Equal(new[] { "TaxRate =0.2 (workbook)" }, lines);
    }

    [Fact]
    public void A_custom_function()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");

        // snippet: custom-function
        workbook.Functions.Add("VAT", args =>
        {
            CellValue amount = args.Number(0);   // converted as Excel converts: "100" is 100
            if (amount.IsError)
                return amount;                  // pass #VALUE! and other errors on
            return amount.AsNumber() * 0.2;
        }, new FunctionOptions { MinArguments = 1, MaxArguments = 1, ScalarParameters = [0] });

        sheet["A1"].Value = 100;
        sheet["A2"].Value = 250;
        sheet["B1"].Formula = "=VAT(A1)";      // 20
        sheet["C1"].Formula = "=VAT(A1:A2)";   // a scalar parameter given a range: spills 20 and 50
        workbook.Recalculate();
        // end-snippet

        Assert.Equal(20, sheet["B1"].Value.AsNumber());
        Assert.Equal(50, sheet["C2"].Value.AsNumber());
    }

    [Fact]
    public void A_custom_function_over_a_range()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");
        sheet["A1"].Value = 3;
        sheet["A2"].Value = "skip";
        sheet["A3"].Value = 4;

        // snippet: custom-function-range
        workbook.Functions.Add("SUMSQUARES", args =>
        {
            double total = 0;
            foreach (CellValue value in args[0].Kind == CellValueKind.Array ? args[0].AsArray() : new[,] { { args[0] } })
            {
                if (value.Kind == CellValueKind.Number)
                    total += value.AsNumber() * value.AsNumber();
            }

            return total;
        }, new FunctionOptions { MinArguments = 1, MaxArguments = 1 });

        sheet["B1"].Formula = "=SUMSQUARES(A1:A3)";   // the range arrives as an array: 25
        // end-snippet
        workbook.Recalculate();

        Assert.Equal(25, sheet["B1"].Value.AsNumber());
    }

    [Fact]
    public void Changed_cells()
    {
        var (workbook, sheet) = Budget();
        var sent = new List<string>();

        // snippet: changed-cells
        sheet["B2"].Value = 1350;
        RecalculationResult result = workbook.Recalculate();
        foreach (Cell cell in result.ChangedCells)   // C2, F2, F3, B6, C6: formulas whose value changed
            sent.Add($"{cell.Worksheet.Name}!{cell.Address} = {cell.Value}");
        // end-snippet

        Assert.Equal(new[] { "Budget!C2 = 16200", "Budget!F2 = 1040", "Budget!F3 = 0.45", "Budget!B6 = 1960", "Budget!C6 = 23520" }, sent);
    }

    [Fact]
    public void Precedents_and_dependents()
    {
        var (_, sheet) = Budget();

        // snippet: precedents
        IReadOnlyList<CellRange> inputs = sheet["F2"].Precedents;   // Budget!F1, Budget!B6
        IReadOnlyList<Cell> readers = sheet["B2"].Dependents;       // Budget!C2, Budget!F3, Budget!B6
        // end-snippet

        Assert.Equal(new[] { "Budget!F1", "Budget!B6" }, inputs.Select(r => r.ToString()));
        Assert.Equal(new[] { "Budget!C2", "Budget!F3", "Budget!B6" }, readers.Select(c => c.ToString()));
        Assert.Equal(new[] { "Budget!B6", "Budget!C2", "Budget!C6", "Budget!F2", "Budget!F3" }, Affected(sheet["B2"]).Select(c => c.ToString()).Order());
    }

    // snippet: affected
    // Every formula an input reaches, directly or through other formulas.
    static List<Cell> Affected(Cell input)
    {
        var seen = new HashSet<string>();
        var found = new List<Cell>();
        var queue = new Queue<Cell>(input.Dependents);
        while (queue.Count > 0)
        {
            Cell cell = queue.Dequeue();
            if (!seen.Add(cell.ToString()))
                continue;
            found.Add(cell);
            foreach (Cell next in cell.Dependents)
                queue.Enqueue(next);
        }

        return found;
    }
    // end-snippet

    [Fact]
    public void Iterative_calculation()
    {
        // snippet: iterative
        var workbook = new Workbook { Iteration = new IterationSettings(Enabled: true) };
        Worksheet loan = workbook.AddSheet("Loan");
        loan["A1"].Value = 1000;                          // opening balance
        loan["A2"].Formula = "=A1+A3";                    // closing balance
        loan["A3"].Formula = "=(A1+A2)/2*0.05";           // interest on the average balance: a circular reference
        workbook.Recalculate();
        // loan["A2"].Value is about 1051.28; without iteration A2 and A3 would be 0
        // end-snippet

        Assert.Equal(1051.28, loan["A2"].Value.AsNumber(), 2);
        Assert.Empty(workbook.Diagnostics);
    }

    [Fact]
    public void Clone_a_template_per_request()
    {
        double[] amounts = [1350, 1500];
        var results = new System.Collections.Concurrent.ConcurrentBag<double>();

        // snippet: clone-template
        // Once, at startup: load and calculate the template.
        Workbook template = XlsxReader.Load(SampleTests.BudgetPath);
        template.Recalculate();

        // Per request, on any thread: a copy costs no file reading and no full calculation.
        Parallel.ForEach(amounts, amount =>
        {
            Workbook workbook = template.Clone();
            workbook["Budget"]["B2"].Value = amount;
            workbook.Recalculate();   // only what depends on B2
            results.Add(workbook["Budget"]["F2"].Value.AsNumber());
        });
        // end-snippet

        Assert.Equal(new[] { 890.0, 1040.0 }, results.Order());
        Assert.Equal(1190, template["Budget"]["F2"].Value.AsNumber());
    }
}
