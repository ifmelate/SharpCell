using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Functions;

/// <summary>T.TEST, F.TEST, CHISQ.TEST, Z.TEST and PROB; expected values are Excel's.</summary>
public class DistributionTestFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public DistributionTestFunctionTests()
    {
        _s = _wb.AddSheet("S");
        // Samples with text, logical values and blanks mixed in (from Excel's T.TEST workbook).
        object[] a = [1, 2, 3, 4, -3, true, "hello"];
        object[] b = [2, -3.3, 0, true, "no", 2, 2];
        for (var i = 0; i < a.Length; i++)
        {
            Set($"A{i + 1}", a[i]);
            Set($"B{i + 1}", b[i]);
        }

        // F.TEST samples.
        double[] c = [1, 2, 3.3, 4.5, -1];
        double[] d = [5, 6, 7, 2.4, -2];
        for (var i = 0; i < c.Length; i++)
        {
            _s[$"C{i + 1}"].Value = c[i];
            _s[$"D{i + 1}"].Value = d[i];
        }

        // CHISQ.TEST observed and expected.
        double[] observed = [1, 2, 3, 4];
        double[] expected = [3, 4, 2, 3];
        for (var i = 0; i < observed.Length; i++)
        {
            _s[$"E{i + 1}"].Value = observed[i];
            _s[$"F{i + 1}"].Value = expected[i];
        }

        // Z.TEST sample and PROB table.
        object[] z = [1, 2, 3, 2, 2.2, 1.8, 2, true, 1, 2];
        for (var i = 0; i < z.Length; i++)
            Set($"G{i + 1}", z[i]);
        double[] values = [1, 2, 3, 4, 5, 6];
        double[] probabilities = [0.2, 0.111, 0.05, 0.07, 0.5, 0.069];
        for (var i = 0; i < values.Length; i++)
        {
            _s[$"H{i + 1}"].Value = values[i];
            _s[$"I{i + 1}"].Value = probabilities[i];
        }
    }

    private void Set(string address, object value)
    {
        _s[address].Value = value switch
        {
            int i => i,
            double x => x,
            bool flag => flag,
            string text => text,
            _ => throw new System.ArgumentException(nameof(value)),
        };
    }

    [Theory]
    [InlineData("=T.TEST(A1:A7,B1:B7,1,1)", 0.15855028798217946)]
    [InlineData("=T.TEST(A1:A7,B1:B7,1,2)", 0.3017830576436469)]
    [InlineData("=T.TEST(A1:A7,B1:B7,1,3)", 0.3019520078734871)]
    [InlineData("=T.TEST(A1:A7,B1:B7,2,1)", 0.3171005759643589)]
    [InlineData("=T.TEST(A1:A7,B1:B7,2,2)", 0.6035661152872938)]
    [InlineData("=T.TEST(A1:A7,B1:B7,2.2,3.3)", 0.6039040157469742)]
    [InlineData("=TTEST(A1:A7,B1:B7,2,3)", 0.6039040157469742)]
    [InlineData("=T.TEST(A1:A7,B1:B7,0.6,1)", "#NUM!")]
    [InlineData("=T.TEST(A1:A7,B1:B7,1,4)", "#NUM!")]
    [InlineData("=T.TEST(A1:A7,B1:B6,1,1)", "#N/A")]
    [InlineData("=T.TEST(A1:A1,B1:B1,1,2)", "#DIV/0!")]
    [InlineData("=F.TEST(C1:C5,D1:D5)", 0.3266664555668095)]
    [InlineData("=F.TEST({1,2,3,4},D1:D4)", 0.5023992107800741)]
    [InlineData("=F.TEST(C1:C3,{2;3;7})", 0.3193277310924369)]
    [InlineData("=FTEST(C1:C5,D1:D5)", 0.3266664555668095)]
    [InlineData("=F.TEST(C1,D1:D5)", "#DIV/0!")]
    [InlineData("=CHISQ.TEST(E1:E4,F1:F4)", 0.3666353840752714)]
    [InlineData("=CHITEST(E1:E4,F1:F4)", 0.3666353840752714)]
    [InlineData("=CHISQ.TEST(E1:E4,F1:F3)", "#N/A")]
    [InlineData("=CHISQ.TEST({1,2;3,4},{2,2;3,3})", 0.3613104285261787)]
    [InlineData("=CHISQ.TEST(E1:E4,{3;0;2;3})", "#DIV/0!")]
    [InlineData("=Z.TEST(G1:G10,2,0.8)", 0.6615388804893105)]
    [InlineData("=Z.TEST(G1:G10,1.7,1)", 0.285470335901444)]
    [InlineData("=Z.TEST(G1:G10,2.3)", 0.9785431996696947)]
    [InlineData("=ZTEST(G1:G10,1.99)", 0.6907332550934567)]
    [InlineData("=Z.TEST(J1:J3,1)", "#N/A")]
    [InlineData("=PROB(H1:H6,I1:I6,2)", 0.111)]
    [InlineData("=PROB(H1:H6,I1:I6,1,3)", 0.361)]
    [InlineData("=PROB(H1:H6,I1:I6,2.4,5.3)", 0.62)]
    [InlineData("=PROB(H1:H6,I1:I6,3,2)", 0)]
    [InlineData("=PROB(H1:H6,I1:I6,\"Hola\")", "#VALUE!")]
    [InlineData("=PROB(H1:H6,I1:I5,1)", "#N/A")]
    [InlineData("=PROB(H1:H5,I1:I5,2)", "#NUM!")]
    [InlineData("=PROB(A7,A7,1)", "#DIV/0!")]
    [InlineData("=PROB({1,2},{0.5,1.5},1)", "#NUM!")]
    public void MatchesExcel(string formula, object expected)
    {
        DistributionFunctionTests.AssertExcel(expected, _wb.Evaluate(formula));
    }

    [Fact]
    public void An_error_in_a_sample_is_the_result()
    {
        _s["J1"].Formula = "=1/0";
        _s["J2"].Value = 1;
        _s["J3"].Value = 2;
        _wb.Recalculate();
        DistributionFunctionTests.AssertExcel("#DIV/0!", _wb.Evaluate("=T.TEST(J1:J3,C1:C5,2,2)"));
        DistributionFunctionTests.AssertExcel("#DIV/0!", _wb.Evaluate("=Z.TEST(J1:J3,1)"));
    }

    [Fact]
    public void Samples_are_arrays_even_in_a_formula_from_before_dynamic_arrays()
    {
        // C1:C5*1 would be reduced to one cell if T.TEST's parameters were scalars.
        var loader = new SheetLoader(_s);
        loader.SetFormula(2, 11, FormulaParser.Parse("=T.TEST(C1:C5*1,D1:D5*1,2,2)", new CellAddress(2, 11)), LoadedFormulaKind.Legacy, null, CellValue.Empty);
        loader.Complete();
        _wb.Recalculate();

        var expected = _wb.Evaluate("=T.TEST(C1:C5,D1:D5,2,2)");
        DistributionFunctionTests.AssertExcel(expected.AsNumber(), _s[2, 11].Value);
    }
}
