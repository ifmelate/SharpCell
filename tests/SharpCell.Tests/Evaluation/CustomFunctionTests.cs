using System.Linq;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Evaluation;

public class CustomFunctionTests
{
    private static CellValue N(double value) => CellValue.Number(value);

    private static (Workbook Workbook, Worksheet Sheet) NewSheet()
    {
        var workbook = new Workbook();
        return (workbook, workbook.AddSheet("S"));
    }

    private static CellValue Twice(FunctionArguments args)
    {
        var x = args.Number(0);
        return x.IsError ? x : x.AsNumber() * 2;
    }

    [Fact]
    public void A_registered_function_is_called_by_formulas()
    {
        var (workbook, sheet) = NewSheet();
        workbook.Functions.Add("TWICE", Twice, new FunctionOptions { MinArguments = 1, MaxArguments = 1 });
        sheet["A1"].Value = 21;
        sheet["A2"].Formula = "=twice(A1)+1";
        workbook.Recalculate();

        Assert.Equal(N(43), sheet["A2"].Value);
        Assert.Equal(N(10), workbook.Evaluate("=Twice(5)"));
        Assert.Empty(workbook.Diagnostics);
    }

    [Fact]
    public void Dependents_follow_the_arguments()
    {
        var (workbook, sheet) = NewSheet();
        workbook.Functions.Add("TWICE", Twice);
        sheet["A2"].Formula = "=TWICE(A1)";
        workbook.Recalculate();
        sheet["A1"].Value = 4;
        workbook.Recalculate();

        Assert.Equal(N(8), sheet["A2"].Value);
    }

    [Fact]
    public void Registering_a_function_recalculates_formulas_that_were_name_errors()
    {
        var (workbook, sheet) = NewSheet();
        sheet["A1"].Formula = "=TWICE(2)";
        workbook.Recalculate();
        Assert.Equal(CellValue.Error(ErrorKind.Name), sheet["A1"].Value);

        workbook.Functions.Add("TWICE", Twice);
        workbook.Recalculate();

        Assert.Equal(N(4), sheet["A1"].Value);
        Assert.Empty(workbook.Diagnostics);
    }

    [Fact]
    public void Removing_a_function_makes_its_calls_name_errors_again()
    {
        var (workbook, sheet) = NewSheet();
        workbook.Functions.Add("TWICE", Twice);
        sheet["A1"].Formula = "=TWICE(2)";
        workbook.Recalculate();

        Assert.True(workbook.Functions.Remove("twice"));
        Assert.False(workbook.Functions.Remove("TWICE"));
        workbook.Recalculate();

        Assert.Equal(CellValue.Error(ErrorKind.Name), sheet["A1"].Value);
    }

    [Fact]
    public void A_name_Excel_knows_but_SharpCell_does_not_implement_can_be_registered()
    {
        var (workbook, sheet) = NewSheet();
        sheet["A1"].Formula = "=WEBSERVICE(\"http://x\")";
        workbook.Recalculate();

        workbook.Functions.Add("WEBSERVICE", _ => "stub");
        workbook.Recalculate();

        Assert.Equal(CellValue.Text("stub"), sheet["A1"].Value);
    }

    [Theory]
    [InlineData("SUM")]
    [InlineData("sum")]
    [InlineData("")]
    [InlineData("1ABC")]
    [InlineData("MY FUNC")]
    [InlineData("MY-FUNC")]
    [InlineData("_xlfn.TWICE")]
    public void Names_that_are_built_in_or_not_function_names_are_rejected(string name)
    {
        var (workbook, _) = NewSheet();
        Assert.Throws<ArgumentException>(() => workbook.Functions.Add(name, Twice));
    }

    [Fact]
    public void Adding_a_name_twice_is_rejected()
    {
        var (workbook, _) = NewSheet();
        workbook.Functions.Add("TWICE", Twice);
        Assert.Throws<ArgumentException>(() => workbook.Functions.Add("twice", Twice));
    }

    [Fact]
    public void The_collection_lists_its_functions()
    {
        var (workbook, _) = NewSheet();
        workbook.Functions.Add("Twice", Twice);
        workbook.Functions.Add("My.Rate_2", _ => 1);

        Assert.True(workbook.Functions.Contains("TWICE"));
        Assert.False(workbook.Functions.Contains("SUM"));
        Assert.Equal(new[] { "My.Rate_2", "Twice" }, workbook.Functions.Names.Order());
    }

    [Theory]
    [InlineData(-1, 1)]
    [InlineData(2, 1)]
    [InlineData(0, 256)]
    public void Options_with_impossible_argument_counts_are_rejected(int min, int max)
    {
        var (workbook, _) = NewSheet();
        Assert.Throws<ArgumentException>(() =>
            workbook.Functions.Add("F", Twice, new FunctionOptions { MinArguments = min, MaxArguments = max }));
    }

    [Fact]
    public void A_scalar_parameter_out_of_range_is_rejected()
    {
        var (workbook, _) = NewSheet();
        Assert.Throws<ArgumentException>(() =>
            workbook.Functions.Add("F", Twice, new FunctionOptions { MaxArguments = 2, ScalarParameters = [2] }));
    }

    [Fact]
    public void A_call_with_the_wrong_number_of_arguments_is_a_value_error()
    {
        var (workbook, _) = NewSheet();
        workbook.Functions.Add("TWICE", Twice, new FunctionOptions { MinArguments = 1, MaxArguments = 1 });

        Assert.Equal(CellValue.Error(ErrorKind.Value), workbook.Evaluate("=TWICE()"));
        Assert.Equal(CellValue.Error(ErrorKind.Value), workbook.Evaluate("=TWICE(1,2)"));
    }

    [Fact]
    public void A_range_argument_arrives_as_an_array_and_a_single_cell_as_its_value()
    {
        var (workbook, sheet) = NewSheet();
        sheet["A1"].Value = 1;
        sheet["A2"].Value = "x";
        CellValue seen = default, single = default;
        workbook.Functions.Add("PEEK", args => { seen = args[0]; single = args[1]; return 0; });
        workbook.Evaluate("=PEEK(S!A1:A3, S!A2)");

        Assert.Equal(CellValueKind.Array, seen.Kind);
        Assert.Equal(new CellValue[,] { { 1 }, { "x" }, { CellValue.Empty } }, seen.AsArray());
        Assert.Equal(CellValue.Text("x"), single);
    }

    [Fact]
    public void A_scalar_parameter_given_a_range_runs_per_element_and_spills()
    {
        var (workbook, sheet) = NewSheet();
        workbook.Functions.Add("TWICE", Twice, new FunctionOptions { MaxArguments = 1, ScalarParameters = [0] });
        sheet["A1"].Value = 1;
        sheet["A2"].Value = 2;
        sheet["B1"].Formula = "=TWICE(A1:A2)";
        workbook.Recalculate();

        Assert.Equal(new[] { N(2), N(4) }, new[] { sheet["B1"].Value, sheet["B2"].Value });
    }

    [Fact]
    public void An_omitted_argument_is_missing_and_errors_arrive_as_values()
    {
        var (workbook, _) = NewSheet();
        var log = new List<string>();
        workbook.Functions.Add("PEEK", args =>
        {
            for (var i = 0; i < args.Count; i++)
                log.Add($"{args[i].Kind}:{args.IsMissing(i)}");
            return 0;
        });
        workbook.Evaluate("=PEEK(1,,1/0)");

        Assert.Equal(new[] { "Number:False", "Missing:True", "Error:False" }, log);
    }

    [Fact]
    public void Typed_readers_coerce_as_Excel_does()
    {
        var (workbook, _) = NewSheet();
        workbook.Functions.Add("KINDS", args =>
            $"{args.Number(0).AsNumber()}|{args.Text(1).AsText()}|{args.Boolean(2).AsBoolean()}|{args.Number(3).AsError().ToText()}");

        Assert.Equal(CellValue.Text("2|1|True|#VALUE!"), workbook.Evaluate("=KINDS(\"2\", 1, 1, \"abc\")"));
    }

    [Fact]
    public void An_array_result_spills()
    {
        var (workbook, sheet) = NewSheet();
        workbook.Functions.Add("PAIR", _ => CellValue.Array(new CellValue[,] { { 1, 2 } }));
        sheet["A1"].Formula = "=PAIR()";
        workbook.Recalculate();

        Assert.Equal(N(2), sheet["B1"].Value);
    }

    [Fact]
    public void A_missing_result_is_a_value_error()
    {
        var (workbook, _) = NewSheet();
        workbook.Functions.Add("NOTHING", _ => CellValue.Missing);
        Assert.Equal(CellValue.Error(ErrorKind.Value), workbook.Evaluate("=NOTHING()"));
    }

    [Fact]
    public void An_exception_is_a_value_error_with_a_diagnostic()
    {
        var (workbook, sheet) = NewSheet();
        workbook.Functions.Add("BOOM", _ => throw new InvalidOperationException("no rate for today"));
        sheet["A1"].Formula = "=BOOM()";
        workbook.Recalculate();

        Assert.Equal(CellValue.Error(ErrorKind.Value), sheet["A1"].Value);
        var diagnostic = Assert.Single(workbook.Diagnostics);
        Assert.Equal(DiagnosticKind.CustomFunctionFailure, diagnostic.Kind);
        Assert.Equal("A1", diagnostic.Address);
        Assert.Contains("BOOM", diagnostic.Message);
        Assert.Contains("no rate for today", diagnostic.Message);
    }

    [Fact]
    public void Caller_is_the_formula_cell_or_null_for_Evaluate()
    {
        var (workbook, sheet) = NewSheet();
        workbook.Functions.Add("WHERE", args => args.Caller?.ToString() ?? "none");
        sheet["C3"].Formula = "=WHERE()";
        workbook.Recalculate();

        Assert.Equal(CellValue.Text("S!C3"), sheet["C3"].Value);
        Assert.Equal(CellValue.Text("none"), workbook.Evaluate("=WHERE()"));
    }

    [Fact]
    public void Arguments_know_the_workbook_settings()
    {
        var workbook = new Workbook { DateSystem = DateSystem.Date1904, Culture = System.Globalization.CultureInfo.GetCultureInfo("de-DE") };
        workbook.Functions.Add("SETTINGS", args => $"{args.DateSystem} {args.Culture.Name}");

        Assert.Equal(CellValue.Text("Date1904 de-DE"), workbook.Evaluate("=SETTINGS()"));
    }

    [Fact]
    public void A_volatile_function_runs_on_every_recalculation()
    {
        var (workbook, sheet) = NewSheet();
        var calls = 0;
        workbook.Functions.Add("TICK", _ => ++calls, new FunctionOptions { IsVolatile = true });
        workbook.Functions.Add("STILL", _ => 1);
        sheet["A1"].Formula = "=TICK()";
        sheet["A2"].Formula = "=STILL()";
        workbook.Recalculate();
        workbook.Recalculate();

        Assert.Equal(N(2), sheet["A1"].Value);
    }

    [Fact]
    public void Changing_the_workbook_from_a_function_is_refused()
    {
        var (workbook, sheet) = NewSheet();
        workbook.Functions.Add("MEDDLE", _ => { sheet["Z1"].Value = 1; return 0; });
        workbook.Functions.Add("RECALC", _ => { workbook.Recalculate(); return 0; });
        workbook.Functions.Add("NAME", _ => { workbook.DefineName("X", "1"); return 0; });
        sheet["A1"].Formula = "=MEDDLE()";
        sheet["A2"].Formula = "=RECALC()";
        sheet["A3"].Formula = "=NAME()";
        workbook.Recalculate();

        Assert.All(new[] { "A1", "A2", "A3" }, a => Assert.Equal(CellValue.Error(ErrorKind.Value), sheet[a].Value));
        Assert.Equal(CellValue.Empty, sheet["Z1"].Value);
        Assert.Equal(3, workbook.Diagnostics.Count(d => d.Kind == DiagnosticKind.CustomFunctionFailure && d.Message.Contains("InvalidOperationException")));

        // After the calculation, changes work again.
        sheet["Z1"].Value = 2;
        Assert.Equal(N(2), sheet["Z1"].Value);
    }

    [Fact]
    public void Arguments_cannot_be_used_after_the_call()
    {
        var (workbook, _) = NewSheet();
        FunctionArguments? kept = null;
        workbook.Functions.Add("KEEP", args => { kept = args; return 0; });
        workbook.Evaluate("=KEEP(1)");

        Assert.Throws<InvalidOperationException>(() => kept![0]);
    }

    [Theory]
    [InlineData("=_xll.TWICE(2)")]
    [InlineData("=_xludf.TWICE(2)")]
    public void Calls_with_add_in_prefixes_find_the_function(string formula)
    {
        var (workbook, sheet) = NewSheet();
        sheet["A1"].Formula = formula;
        workbook.Recalculate();
        workbook.Functions.Add("TWICE", Twice);
        workbook.Recalculate();

        Assert.Equal(N(4), sheet["A1"].Value);
        Assert.Equal(formula, sheet["A1"].Formula);
    }

    [Fact]
    public void A_workbook_using_a_custom_function_saves_without_keeping_uncalculated_results()
    {
        using var file = new Xlsx.TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>2</v></c><c r=\"B1\"><f>MYTAX(A1)</f><v>0</v></c></row>").Build();
        var workbook = XlsxReader.Load(file);
        workbook.Functions.Add("MYTAX", args => args.Number(0).AsNumber() / 2);
        workbook.Recalculate();

        using var saved = new MemoryStream();
        XlsxWriter.Save(workbook, saved);
        saved.Position = 0;

        Assert.Equal(N(1), XlsxReader.Load(saved)["S"]["B1"].Value);
    }

    [Fact]
    public void Each_workbook_has_its_own_functions()
    {
        var (first, _) = NewSheet();
        var (second, _) = NewSheet();
        first.Functions.Add("TWICE", Twice);

        Assert.Equal(CellValue.Error(ErrorKind.Name), second.Evaluate("=TWICE(1)"));
    }

    [Fact]
    public void Adding_a_table_from_a_function_is_refused_before_anything_is_calculated()
    {
        var (workbook, sheet) = NewSheet();
        Exception? caught = null;
        workbook.Functions.Add("HOOK", _ =>
        {
            try { sheet.AddTable("T", "A1:A2"); } catch (Exception ex) { caught = ex; }
            return 5;
        });
        sheet["A1"].Formula = "=C1+1";
        sheet["C1"].Formula = "=HOOK()";
        workbook.Recalculate();

        Assert.IsType<InvalidOperationException>(caught);
        Assert.Equal(N(6), sheet["A1"].Value);
        Assert.Empty(workbook.Diagnostics);
    }

    [Fact]
    public void An_array_argument_is_a_copy_the_function_may_change()
    {
        var template = new Workbook();
        var sheet = template.AddSheet("S");
        sheet["A1"].Formula = "=LET(x, {1,2}, IFERROR(MUT(x), 0) + SUM(x))";
        template.Recalculate();
        var clone = template.Clone();
        clone.Functions.Add("MUT", a => { a[0].AsArray()[0, 0] = CellValue.Number(99); return 0; });
        clone.Recalculate();

        template.Culture = System.Globalization.CultureInfo.InvariantCulture;   // calculates again without parsing again
        template.Recalculate();

        Assert.Equal(N(3), sheet["A1"].Value);
        Assert.Equal(N(3), clone["S"]["A1"].Value);
    }

    [Fact]
    public void A_function_sees_the_cancellation_token_of_the_calculation()
    {
        var (workbook, sheet) = NewSheet();
        using var cancel = new CancellationTokenSource();
        workbook.Functions.Add("SLOW", args =>
        {
            cancel.Cancel();
            args.CancellationToken.ThrowIfCancellationRequested();
            return 1;
        });
        sheet["A1"].Formula = "=SLOW()";

        Assert.ThrowsAny<OperationCanceledException>(() => workbook.Recalculate(cancel.Token));
        Assert.Empty(workbook.Diagnostics);
    }
}
