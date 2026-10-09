using System;
using System.Globalization;
using System.Linq;
using SharpCell.Conformance;

namespace SharpCell.Tests.Xlsx;

public class CorpusOverridesTests
{
    private const string Irr = "ironcalc/calc_tests/FINANCIAL/IRR.xlsx";

    private const string Json = """
        [
          { "file": "a.xlsx", "culture": "en-IE", "reason": "Calculated with day-first dates." },
          { "file": "b.xlsx", "cells": ["Sheet 1!B2:C3", "Other!A1"], "absolute": 1e-6, "reason": "Solver noise." },
          { "file": "b.xlsx", "cells": ["Sheet 1!D9"], "relative": 1e-8, "reason": "Documented accuracy." }
        ]
        """;

    [Fact]
    public void Finds_the_rule_for_a_cell()
    {
        var overrides = CorpusOverrides.Parse(Json);

        Assert.Equal(CultureInfo.GetCultureInfo("en-IE"), overrides.CultureOf("a.xlsx"));
        Assert.Null(overrides.CultureOf("b.xlsx"));
        Assert.Equal(1e-6, overrides.For("b.xlsx", "Sheet 1", new CellAddress(3, 3))!.Absolute);
        Assert.Equal(1e-6, overrides.For("b.xlsx", "Other", new CellAddress(1, 1))!.Absolute);
        Assert.Equal(1e-8, overrides.For("b.xlsx", "Sheet 1", new CellAddress(9, 4))!.Relative);
        Assert.Null(overrides.For("b.xlsx", "Sheet 1", new CellAddress(4, 2)));
        Assert.Null(overrides.For("a.xlsx", "Sheet 1", new CellAddress(2, 2)));
    }

    [Theory]
    [InlineData("""[{ "file": "a.xlsx", "culture": "en-IE" }]""")]
    [InlineData("""[{ "file": "a.xlsx", "reason": "Nothing to do." }]""")]
    [InlineData("""[{ "file": "a.xlsx", "cells": ["A1"], "absolute": 1e-6, "reason": "No sheet." }]""")]
    [InlineData("""[{ "file": "a.xlsx", "cells": ["S!A1"], "reason": "No tolerance." }]""")]
    public void Every_rule_has_a_reason_and_an_effect(string json)
    {
        Assert.Throws<FormatException>(() => CorpusOverrides.Parse(json));
    }

    [Fact]
    public void Stated_tolerance_makes_noise_match_and_is_counted_apart()
    {
        // NPV at the IRR in IRR.xlsx: Excel and SharpCell both give zero up to the solver's accuracy.
        var plain = CorpusRunner.Run(CorpusFiles.PathOf(Irr), Irr, TimeSpan.FromMinutes(1));
        var noisy = plain.Cells.Where(c => !c.Passed).ToList();
        Assert.Contains(noisy, c => c.Sheet == "General" && c.Address == "M2");

        var overrides = CorpusOverrides.Parse("""
            [{ "file": "ironcalc/calc_tests/FINANCIAL/IRR.xlsx", "cells": ["General!M2:M19"], "absolute": 1e-6, "reason": "NPV at the IRR." }]
            """);
        var widened = CorpusRunner.Run(CorpusFiles.PathOf(Irr), Irr, TimeSpan.FromMinutes(1), overrides);

        var m2 = widened.Cells.Single(c => c.Sheet == "General" && c.Address == "M2");
        Assert.True(m2.Passed);
        Assert.True(m2.Widened);
        Assert.Equal(plain.Passed + noisy.Count(c => c.Sheet == "General" && c.Cell.Column == 13), widened.Passed);
        Assert.DoesNotContain(widened.Cells, c => c.Widened && c.Sheet != "General");
    }

    [Fact]
    public void Culture_of_a_file_is_used_to_calculate_it()
    {
        const string file = "ironcalc/calc_tests/TEXT/T_VALUE_VALUETOTEXT.xlsx";
        var overrides = CorpusOverrides.Parse("""
            [{ "file": "ironcalc/calc_tests/TEXT/T_VALUE_VALUETOTEXT.xlsx", "culture": "en-IE", "reason": "Euro and day-first dates." }]
            """);
        var result = CorpusRunner.Run(CorpusFiles.PathOf(file), file, TimeSpan.FromMinutes(1), overrides);
        Assert.All(result.Cells, c => Assert.True(c.Passed, $"{c.Sheet}!{c.Address}: Excel {c.Expected}, SharpCell {c.Actual}"));
    }

    [Fact]
    public void Report_lists_the_rules_and_counts_widened_cells()
    {
        var overrides = CorpusOverrides.Parse("""
            [{ "file": "ironcalc/calc_tests/FINANCIAL/IRR.xlsx", "cells": ["General!M2:M19"], "absolute": 1e-6, "reason": "NPV at the IRR." }]
            """);
        var result = CorpusRunner.Run(CorpusFiles.PathOf(Irr), Irr, TimeSpan.FromMinutes(1), overrides);
        var report = new Report([result], overrides);

        var widened = result.Cells.Count(c => c.Widened);
        Assert.True(widened > 0);
        Assert.Equal(widened, report.Widened);
        Assert.Contains("NPV at the IRR.", report.ToMarkdown());
        Assert.Contains($"\"widened\": {widened}", report.ToJson());
    }

    [Fact]
    public void Headline_does_not_round_up_to_100_percent()
    {
        var cells = Enumerable.Range(1, 76082)
            .Select(row => new CellOutcome("S", new CellAddress(row, 1), "=1", CellValue.Number(1), CellValue.Number(1), row > 2, []))
            .ToList();
        var markdown = new Report([new FileResult("f.xlsx", cells, 0, null)]).ToMarkdown();
        Assert.Contains("**76080 of 76082 cells match Excel (99.997%).**", markdown);
    }
}
