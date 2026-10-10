using System;
using System.Linq;
using SharpCell.Conformance;

namespace SharpCell.Tests.Xlsx;

/// <summary>Corpus files that must match Excel in every cell, not only in the cells the report recorded.</summary>
public class ExactCorpusTests
{
    public static TheoryData<string> Files() =>
    [
        "excel/tables.xlsx",
        "excel-web/tables-spike.xlsx",
        "excel-web/hidden-rows.xlsx",
        "excel-web/filter-mode.xlsx",
        "excel-web/spill-blocked.xlsx",
    ];

    [Theory]
    [MemberData(nameof(Files))]
    public void Every_cell_matches_Excel(string file)
    {
        var result = CorpusRunner.Run(CorpusFiles.PathOf(file), file, TimeSpan.FromMinutes(2), CorpusOverrides.Load(CorpusFiles.Root));
        Assert.Null(result.Error);
        Assert.NotEmpty(result.Cells);
        var failing = result.Cells.Where(c => !c.Passed)
            .Select(c => $"{c.Sheet}!{c.Address} {c.Formula ?? "(spilled)"}: Excel {c.Expected}, SharpCell {c.Actual}")
            .ToList();
        Assert.True(failing.Count == 0, $"{file}:\n" + string.Join("\n", failing));
    }
}
