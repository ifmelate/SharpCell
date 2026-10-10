using System.IO.Compression;
using System.Text.RegularExpressions;
using SharpCell;
using SharpCell.Parsing;
using SharpCell.Tests.Xlsx;

namespace SharpCell.Tests.Parsing;

/// <summary>Formulas printed the way .xlsx files store them, so Excel reads them as written.</summary>
public partial class FileSyntaxTests
{
    private static readonly CellAddress Origin = new(3, 3);

    private static string ForFile(string text, Func<CellAddress, string?>? tableAt = null)
    {
        var node = FormulaParser.Parse(text, Origin);
        var printed = FormulaPrinter.PrintForFile(node, Origin, tableAt ?? (_ => null));

        // Reading the printed text back gives the same formula.
        Assert.Equal(printed, FormulaPrinter.PrintForFile(FormulaParser.Parse(printed, Origin), Origin, tableAt ?? (_ => null)));
        return printed;
    }

    [Theory]
    [InlineData("XLOOKUP(A1,B:B,C:C)", "_xlfn.XLOOKUP(A1,B:B,C:C)")]
    [InlineData("FILTER(A1:A9,B1:B9>1)", "_xlfn._xlws.FILTER(A1:A9,B1:B9>1)")]
    [InlineData("SORT(A1:A3)", "_xlfn._xlws.SORT(A1:A3)")]
    [InlineData("SUM(A1:A3)", "SUM(A1:A3)")]
    [InlineData("STDEV.S(A1:A3)", "_xlfn.STDEV.S(A1:A3)")]
    [InlineData("IFERROR(1/0,CONCAT(\"a\",\"b\"))", "IFERROR(1/0,_xlfn.CONCAT(\"a\",\"b\"))")]
    [InlineData("LET(x,1,x+1)", "_xlfn.LET(_xlpm.x,1,_xlpm.x+1)")]
    [InlineData("LAMBDA(a,b,a*b)(2,3)", "_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a*_xlpm.b)(2,3)")]
    [InlineData("LET(x,1,y,x+1,y*x)", "_xlfn.LET(_xlpm.x,1,_xlpm.y,_xlpm.x+1,_xlpm.y*_xlpm.x)")]
    [InlineData("LET(f,LAMBDA(x,x+1),f(2))", "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),_xlpm.F(2))")]
    [InlineData("x+LET(x,1,x)", "x+_xlfn.LET(_xlpm.x,1,_xlpm.x)")]
    [InlineData("LET(x,x+1,x)", "_xlfn.LET(_xlpm.x,x+1,_xlpm.x)")]
    [InlineData("LAMBDA(a,[b],a)(1)", "_xlfn.LAMBDA(_xlpm.a,_xlop.b,_xlpm.a)(1)")]
    public void Newer_functions_get_their_file_prefixes(string formula, string expected) =>
        Assert.Equal(expected, ForFile(formula));

    [Theory]
    [InlineData("A1#", "_xlfn.ANCHORARRAY(A1)")]
    [InlineData("SUM(Sheet2!B2#)", "SUM(_xlfn.ANCHORARRAY(Sheet2!B2))")]
    [InlineData("@A1:A3", "_xlfn.SINGLE(A1:A3)")]
    [InlineData("@A1:A3*2", "_xlfn.SINGLE(A1:A3)*2")]
    public void Spill_and_implicit_intersection_operators_become_functions(string formula, string expected) =>
        Assert.Equal(expected, ForFile(formula));

    [Theory]
    [InlineData("Sales[@Units]", "Sales[[#This Row],[Units]]")]
    [InlineData("Sales[@[Unit Price]]*2", "Sales[[#This Row],[Unit Price]]*2")]
    [InlineData("SUM(Sales[@[Units]:[Price]])", "SUM(Sales[[#This Row],[Units]:[Price]])")]
    [InlineData("Sales[Units]", "Sales[Units]")]
    [InlineData("Sales[[#Headers],[Units]]", "Sales[[#Headers],[Units]]")]
    [InlineData("Sales[#All]", "Sales[#All]")]
    [InlineData("Sales[[Units]:[Price]]", "Sales[[Units]:[Price]]")]
    public void Structured_references_are_written_without_the_at_sign(string formula, string expected) =>
        Assert.Equal(expected, ForFile(formula));

    [Theory]
    [InlineData("[@Units]", "Sales[[#This Row],[Units]]")]
    [InlineData("[Units]", "Sales[Units]")]
    [InlineData("[@['#Items]]", "Sales[[#This Row],['#Items]]")]
    public void A_reference_inside_a_table_gets_the_table_name(string formula, string expected) =>
        Assert.Equal(expected, ForFile(formula, cell => cell == Origin ? "Sales" : null));

    [Fact]
    public void A_table_reference_outside_any_table_stays_as_written() =>
        Assert.Equal("[@Units]", ForFile("[@Units]"));

    [GeneratedRegex(@"_xlfn\.(_xlws\.)?([A-Z][A-Z0-9.]*)\(")]
    private static partial Regex PrefixedCall();

    [Fact]
    public void Every_prefixed_function_in_the_corpus_is_known()
    {
        var found = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var file in CorpusFiles.All)
        {
            using var zip = ZipFile.OpenRead(CorpusFiles.PathOf(file));
            foreach (var entry in zip.Entries.Where(e => e.FullName.EndsWith(".xml", StringComparison.Ordinal)))
            {
                using var reader = new StreamReader(entry.Open());
                foreach (Match match in PrefixedCall().Matches(reader.ReadToEnd()))
                    found[match.Groups[2].Value] = match.Groups[1].Success ? "_xlfn._xlws." : "_xlfn.";
            }
        }

        Assert.NotEmpty(found);
        Assert.All(found, pair => Assert.Equal((pair.Key, pair.Value), (pair.Key, FutureFunctions.Prefix(pair.Key))));
    }
}
