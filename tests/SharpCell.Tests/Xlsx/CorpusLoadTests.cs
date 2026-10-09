using System.Linq;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class CorpusLoadTests
{
    public static TheoryData<string> Files() => [.. CorpusFiles.All];

    [Fact]
    public void Corpus_is_present()
    {
        Assert.True(CorpusFiles.All.Count >= 239, $"Expected the corpus under {CorpusFiles.Root}.");
    }

    [Theory]
    [MemberData(nameof(Files))]
    public void Every_corpus_file_loads(string file)
    {
        var workbook = XlsxReader.Load(CorpusFiles.PathOf(file));

        Assert.NotEmpty(workbook.Sheets);
    }
}

public class CorpusParseTests
{
    // Formulas of the corpus SharpCell cannot parse, by file. Anything new here is a parser regression.
    private static readonly string[] Expected = [];

    [Fact]
    public void Corpus_formulas_parse_except_known_gaps()
    {
        var unsupported = new System.Collections.Generic.List<string>();
        foreach (var file in CorpusFiles.All)
        {
            var workbook = XlsxReader.Load(CorpusFiles.PathOf(file));
            foreach (var sheet in workbook.Sheets)
            {
                foreach (var cell in sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
                {
                    if (cell.Data.Formula is SharpCell.Parsing.UnsupportedNode u)
                        unsupported.Add($"{file} {sheet.Name}!{new CellAddress(cell.Row, cell.Column)} {u.Text} :: {u.Reason}");
                }
            }
        }

        Assert.Equal(Expected, unsupported);
    }
}
