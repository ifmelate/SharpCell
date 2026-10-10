using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using SharpCell.Tests.Xlsx;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Model;

public class CloneTests
{
    private static CellValue N(double value) => CellValue.Number(value);

    private static (Workbook Workbook, Worksheet Sheet) Template()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = 2;
        sheet["A2"].Formula = "=A1*10";
        sheet["B1"].Formula = "=SEQUENCE(A1)";
        workbook.Recalculate();
        return (workbook, sheet);
    }

    private static string[] Snapshot(Workbook workbook) =>
        [.. workbook.Sheets.SelectMany(s => s.Cells).Select(c => $"{c}={c.Value}|{c.Formula}")];

    [Fact]
    public void A_clone_has_the_same_cells_and_needs_no_calculation()
    {
        var (template, _) = Template();
        var clone = template.Clone();

        Assert.Equal(Snapshot(template), Snapshot(clone));
        Assert.False(clone.Calculation.HasDirty);
        Assert.Empty(clone.Recalculate().ChangedCells);
        Assert.Equal(0, clone.Calculation.EvaluationCount);
    }

    [Fact]
    public void A_clone_has_sheets_of_its_own()
    {
        var (template, sheet) = Template();
        var clone = template.Clone();

        Assert.NotSame(sheet, clone["S"]);
        Assert.Same(clone, clone["S"].Workbook);
        Assert.Equal(template.Sheets.Select(s => s.Name), clone.Sheets.Select(s => s.Name));
    }

    [Fact]
    public void Changing_a_clone_leaves_the_template_alone_and_the_other_way_round()
    {
        var (template, sheet) = Template();
        var clone = template.Clone();

        clone["S"]["A1"].Value = 3;
        var changed = clone.Recalculate().ChangedCells.Select(c => c.ToString());

        Assert.Equal(new[] { "S!A2", "S!B3" }, changed);
        Assert.Equal(N(30), clone["S"]["A2"].Value);
        Assert.Equal(N(3), clone["S"]["B3"].Value);
        Assert.Equal(N(20), sheet["A2"].Value);
        Assert.Equal(CellValue.Empty, sheet["B3"].Value);
        Assert.False(template.Calculation.HasDirty);

        sheet["A1"].Value = 1;
        template.Recalculate();
        Assert.Equal(N(30), clone["S"]["A2"].Value);
    }

    [Fact]
    public void Formulas_out_of_date_in_the_template_are_out_of_date_in_the_clone()
    {
        var (template, sheet) = Template();
        sheet["A1"].Value = 4;
        var clone = template.Clone();

        Assert.True(clone.Calculation.HasDirty);
        clone.Recalculate();
        Assert.Equal(N(40), clone["S"]["A2"].Value);
        Assert.Equal(N(20), sheet["A2"].Value);
    }

    [Fact]
    public void A_spill_in_a_clone_reacts_to_blocking_cells()
    {
        var (template, _) = Template();
        var clone = template.Clone();

        clone["S"]["B2"].Value = "x";
        clone.Recalculate();

        Assert.Equal(CellValue.Error(ErrorKind.Spill), clone["S"]["B1"].Value);
        Assert.Equal(N(1), template["S"]["B1"].Value);

        clone["S"]["B2"].Value = CellValue.Empty;
        clone.Recalculate();
        Assert.Equal(N(2), clone["S"]["B2"].Value);
    }

    [Fact]
    public void Precedents_and_dependents_of_a_clone_are_its_own_cells()
    {
        var (template, _) = Template();
        var clone = template.Clone();

        Assert.All(clone["S"]["A1"].Dependents, c => Assert.Same(clone["S"], c.Worksheet));
        Assert.Equal(new[] { "S!B1", "S!A2" }, clone["S"]["A1"].Dependents.Select(c => c.ToString()));
        Assert.Same(clone["S"], clone["S"]["A2"].Precedents.Single().Worksheet);
    }

    [Fact]
    public void Settings_names_tables_hidden_rows_and_filters_are_copied()
    {
        var template = new Workbook { Culture = System.Globalization.CultureInfo.GetCultureInfo("de-DE"), DateSystem = DateSystem.Date1904 };
        var sheet = template.AddSheet("S");
        sheet["A1"].Value = "Units";
        sheet["A2"].Value = 5;
        sheet["A3"].Value = 7;
        sheet.AddTable("Sales", "A1:A3");
        template.DefineName("Rate", "=2");
        template.DefineName("Local", "=S!A2", sheet);
        sheet.SetRowHidden(3, true);
        sheet.FilterMode = true;
        sheet["C1"].Formula = "=SUM(Sales[Units])*Rate+Local";
        sheet["C2"].Formula = "=SUBTOTAL(9,A2:A3)";
        template.Recalculate();

        var clone = template.Clone();
        var copy = clone["S"];

        Assert.Equal("de-DE", clone.Culture.Name);
        Assert.Equal(DateSystem.Date1904, clone.DateSystem);
        Assert.Equal(template.DefinedNames.Select(n => (n.Name, n.Formula, n.Scope?.Name)), clone.DefinedNames.Select(n => (n.Name, n.Formula, n.Scope?.Name)));
        Assert.Same(copy, clone.DefinedNames[1].Scope);
        Assert.True(clone.TryGetTable("sales", out var table));
        Assert.Same(copy, table!.Worksheet);
        Assert.True(copy.IsRowHidden(3));
        Assert.True(copy.FilterMode);

        copy["A2"].Value = 6;
        clone.Recalculate();
        Assert.Equal(N(32), copy["C1"].Value);
        Assert.Equal(N(6), copy["C2"].Value);
        Assert.Equal(N(29), sheet["C1"].Value);

        copy.SetRowHidden(3, false);
        Assert.True(sheet.IsRowHidden(3));
    }

    [Fact]
    public void Diagnostics_are_copied_with_the_clone_s_sheets()
    {
        var template = new Workbook();
        var sheet = template.AddSheet("S");
        sheet["A1"].Formula = "=B1";
        sheet["B1"].Formula = "=A1";
        template.Recalculate();

        var clone = template.Clone();

        Assert.Equal(template.Diagnostics.Select(d => d.ToString()), clone.Diagnostics.Select(d => d.ToString()));
        Assert.All(clone.Diagnostics, d => Assert.Same(clone["S"], d.Sheet));

        clone["S"]["B1"].Value = 1;
        clone.Recalculate();
        Assert.Empty(clone.Diagnostics);
        Assert.NotEmpty(template.Diagnostics);
    }

    [Fact]
    public void Custom_functions_are_copied_and_stay_separate()
    {
        var template = new Workbook();
        template.AddSheet("S")["A1"].Formula = "=TWICE(2)";
        template.Functions.Add("TWICE", args => args.Number(0).AsNumber() * 2);
        template.Recalculate();

        var clone = template.Clone();
        clone.Functions.Add("THRICE", args => args.Number(0).AsNumber() * 3);

        Assert.Equal(N(5), clone.Evaluate("=TWICE(1)+THRICE(1)"));
        Assert.False(template.Functions.Contains("THRICE"));
    }

    [Fact]
    public void Volatile_formulas_stay_volatile_in_a_clone()
    {
        var template = new Workbook();
        var sheet = template.AddSheet("S");
        var calls = 0;
        template.Functions.Add("TICK", _ => ++calls, new FunctionOptions { IsVolatile = true });
        sheet["A1"].Formula = "=TICK()";
        template.Recalculate();

        var clone = template.Clone();
        clone.Recalculate();

        Assert.Equal(N(2), clone["S"]["A1"].Value);
        Assert.Equal(N(1), sheet["A1"].Value);
    }

    [Fact]
    public void Changes_pending_in_the_template_are_not_reported_by_the_clone()
    {
        var (template, sheet) = Template();
        sheet["A1"].Value = 5;
        template.Evaluate("=S!A2");

        var clone = template.Clone();

        Assert.DoesNotContain(clone.Recalculate().ChangedCells, c => c.Address == "A2");
    }

    [Fact]
    public void An_empty_workbook_clones()
    {
        var clone = new Workbook().Clone();
        Assert.Empty(clone.Sheets);
        Assert.Empty(clone.Recalculate().ChangedCells);
    }

    [Fact]
    public void An_array_formula_area_in_a_clone_is_still_protected()
    {
        using var file = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:A2\">{1;2}</f><v>1</v></c></row><row r=\"2\"><c r=\"A2\"><v>2</v></c></row>").Build();
        var clone = XlsxReader.Load(file).Clone();

        Assert.Throws<InvalidOperationException>(() => clone["S"]["A2"].Value = 5);
    }

    [Fact]
    public void A_clone_of_a_loaded_workbook_saves_into_the_same_file()
    {
        var path = CorpusFiles.PathOf("excel/templates/invoice.xlsx");
        var template = XlsxReader.Load(path);
        template.Recalculate();

        var clone = template.Clone();
        Assert.Equal(XlsxWriterTests.Save(template, new XlsxWriteOptions { KeepUncalculated = true }),
            XlsxWriterTests.Save(clone, new XlsxWriteOptions { KeepUncalculated = true }));
    }

    [Fact]
    public void Saving_a_changed_clone_matches_saving_the_same_change_to_the_original()
    {
        using var file = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>2</v></c><c r=\"B1\"><f>A1*2</f><v>4</v></c></row>").Build();
        var original = XlsxReader.Load(file);
        var clone = original.Clone();

        original["S"]["A1"].Value = 5;
        original.Recalculate();
        clone["S"]["A1"].Value = 5;
        clone.Recalculate();

        Assert.Equal(XlsxWriterTests.Save(original), XlsxWriterTests.Save(clone));
    }

    [Fact]
    public void Every_corpus_workbook_clones_into_an_equal_workbook()
    {
        foreach (var relative in CorpusFiles.All)
        {
            Workbook template;
            try
            {
                template = XlsxReader.Load(CorpusFiles.PathOf(relative));
            }
            catch (NotSupportedException)
            {
                continue;
            }

            template.Clock = new FixedClock();

            // Before calculation: saved results and formulas waiting to be calculated.
            var uncalculated = template.Clone();
            Assert.True(Snapshot(template).SequenceEqual(Snapshot(uncalculated)), relative);

            // After calculation: values, spills and diagnostics.
            template.Recalculate();
            var calculated = template.Clone();
            Assert.True(Snapshot(template).SequenceEqual(Snapshot(calculated)), relative);
            Assert.Equal(template.Diagnostics.Select(d => d.ToString()), calculated.Diagnostics.Select(d => d.ToString()));

            // The uncalculated clone calculates what the template did, given the same clock and random numbers.
            var fresh = XlsxReader.Load(CorpusFiles.PathOf(relative));
            fresh.Random = new Random(7);
            fresh.Clock = template.Clock;
            fresh.Recalculate();
            uncalculated.Random = new Random(7);
            uncalculated.Clock = template.Clock;
            uncalculated.Recalculate();
            Assert.True(Snapshot(fresh).SequenceEqual(Snapshot(uncalculated)), relative);
        }
    }

    [Fact]
    public void Many_threads_can_clone_one_template_at_once()
    {
        var template = XlsxReader.Load(CorpusFiles.PathOf("excel/templates/invoice.xlsx"));
        template.Clock = new FixedClock();   // the template has NOW(); clones share the clock
        template.Recalculate();
        var sheet = template.Sheets[0];
        var input = sheet.Cells.First(c => c.Formula is null && c.Value.Kind == CellValueKind.Number && c.Dependents.Count > 0);
        var expected = new ConcurrentDictionary<double, string[]>();
        for (var i = 0; i < 4; i++)
        {
            var copy = template.Clone();
            copy[sheet.Name][input.Address].Value = i;
            copy.Recalculate();
            expected[i] = Snapshot(copy);
        }

        var before = Snapshot(template);
        var failures = new ConcurrentBag<string>();
        Parallel.For(0, 64, new ParallelOptions { MaxDegreeOfParallelism = 8 }, n =>
        {
            var copy = template.Clone();
            copy[sheet.Name][input.Address].Value = n % 4;
            copy.Recalculate();
            if (!Snapshot(copy).SequenceEqual(expected[n % 4]))
                failures.Add($"clone {n} differs");
        });

        Assert.Empty(failures);
        Assert.Equal(before, Snapshot(template));
        Assert.False(template.Calculation.HasDirty);
    }

    private sealed class FixedClock : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => new(2026, 10, 10, 12, 0, 0, TimeSpan.Zero);

        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }
}
