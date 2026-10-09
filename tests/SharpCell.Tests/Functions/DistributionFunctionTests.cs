using System;
using SharpCell;

namespace SharpCell.Tests.Functions;

/// <summary>Continuous distributions and the special functions; expected values are Excel's.</summary>
public class DistributionFunctionTests
{
    private readonly Workbook _wb = new();

    public DistributionFunctionTests()
    {
        _wb.AddSheet("S");
    }

    internal static void AssertExcel(object expected, CellValue actual, double tolerance = 1e-9)
    {
        switch (expected)
        {
            case double number:
                Assert.Equal(CellValueKind.Number, actual.Kind);
                var a = actual.AsNumber();
                Assert.True(a == number || Math.Abs(a - number) <= tolerance * Math.Max(Math.Abs(a), Math.Abs(number)),
                    $"expected {number:R}, got {a:R}");
                break;
            case int whole:
                AssertExcel((double)whole, actual, tolerance);
                break;
            case string error:
                Assert.Equal(error, actual.ToString());
                break;
            default:
                throw new ArgumentException("Unsupported expectation.", nameof(expected));
        }
    }

    [Theory]
    // Normal
    [InlineData("=NORM.DIST(42,40,1.5,TRUE)", 0.9087887802741321)]
    [InlineData("=NORM.DIST(42,40,1.5,FALSE)", 0.10934004978399575)]
    [InlineData("=NORM.DIST(0,40,1.5,TRUE)", 5.7347825020024226e-157)]
    [InlineData("=NORM.DIST(42,40,0,TRUE)", "#NUM!")]
    [InlineData("=NORM.DIST(42,40,1.5,\"abc\")", "#VALUE!")]
    [InlineData("=NORM.DIST(42,40,1.5,\"true\")", 0.9087887802741321)]
    [InlineData("=NORM.DIST(42,40,1.5,6)", 0.9087887802741321)]
    [InlineData("=NORM.DIST(42,40,1.5,)", 0.10934004978399575)]
    [InlineData("=NORMDIST(42,40,1.5,TRUE)", 0.9087887802741321)]
    [InlineData("=NORM.INV(0.9087887802741321,40,1.5)", 42)]
    [InlineData("=NORM.INV(0.908789,40,1.5)", 42.00000200956616)]
    [InlineData("=NORM.INV(1e-7,0,1)", -5.199337582192817)]
    [InlineData("=NORM.INV(0,0,1)", "#NUM!")]
    [InlineData("=NORM.INV(1,0,1)", "#NUM!")]
    [InlineData("=NORM.INV(0.5,0,0)", "#NUM!")]
    [InlineData("=NORMINV(0.025,0,1)", -1.9599639845400538)]
    [InlineData("=NORM.S.DIST(1.333333,TRUE)", 0.9087887256040951)]
    [InlineData("=NORM.S.DIST(1.333333,FALSE)", 0.16401014756936722)]
    [InlineData("=NORM.S.DIST(-10,TRUE)", 7.619853024160475e-24)]
    [InlineData("=NORM.S.DIST(,)", 0.3989422804014327)]
    [InlineData("=NORMSDIST(1)", 0.841344746068543)]
    [InlineData("=NORM.S.INV(0.908789)", 1.3333346730441071)]
    [InlineData("=NORM.S.INV(0.5000003989422803)", 9.99999999800783e-07)]
    [InlineData("=NORM.S.INV(7.619853024160475e-24)", -10.000000000000002)]
    [InlineData("=NORM.S.INV(0.999999)", 4.753424308817089)]
    [InlineData("=NORM.S.INV(TRUE)", "#NUM!")]
    [InlineData("=NORMSINV(0.975)", 1.9599639845400536)]
    [InlineData("=PHI(3.45)", 0.0010382812956614103)]
    [InlineData("=GAUSS(-1.96)", -0.47500210485177957)]
    [InlineData("=GAUSS(1e-7)", 3.9894228032189005e-08)]
    [InlineData("=GAUSS(10)", 0.5)]
    [InlineData("=STANDARDIZE(12.3,12.2,1)", 0.10000000000000142)]
    [InlineData("=STANDARDIZE(1,1,0)", "#NUM!")]
    // Lognormal
    [InlineData("=LOGNORM.DIST(4,3.5,1.2,TRUE)", 0.03908355570680048)]
    [InlineData("=LOGNORM.DIST(4,3.5,1.2,FALSE)", 0.01761759668181922)]
    [InlineData("=LOGNORM.DIST(1e-10,0,1,TRUE)", 1.284175630643432e-117)]
    [InlineData("=LOGNORM.DIST(0,3.5,1.2,TRUE)", "#NUM!")]
    [InlineData("=LOGNORMDIST(4,3.5,1.2)", 0.03908355570680048)]
    [InlineData("=LOGNORM.INV(0.039084,3.5,1.2)", 4.000025218680638)]
    [InlineData("=LOGNORM.INV(0.999999,0,1)", 115.98075925033795)]
    [InlineData("=LOGNORM.INV(0,0,1)", "#NUM!")]
    [InlineData("=LOGINV(0.5,2,1)", 7.38905609893065)]
    // Confidence intervals
    [InlineData("=CONFIDENCE.NORM(0.05,1,50)", 0.2771807648699355)]
    [InlineData("=CONFIDENCE.NORM(0.999999,1,50)", 1.7724538511537302e-07)]
    [InlineData("=CONFIDENCE.NORM(0.05,1,1.999999)", 1.9599639845400536)]
    [InlineData("=CONFIDENCE.NORM(0.05,1,0.999999)", "#NUM!")]
    [InlineData("=CONFIDENCE.NORM(1,1,50)", "#NUM!")]
    [InlineData("=CONFIDENCE.NORM(\"5%\",1,50)", 0.2771807648699355)]
    [InlineData("=CONFIDENCE(0.05,1,50)", 0.2771807648699355)]
    [InlineData("=CONFIDENCE.T(0.05,1,50)", 0.2841968554957298)]
    [InlineData("=CONFIDENCE.T(0.3,0.5,2)", 0.6938875986353249)]
    [InlineData("=CONFIDENCE.T(0.05,1,1)", "#DIV/0!")]
    [InlineData("=CONFIDENCE.T(0.05,0,50)", "#NUM!")]
    // Exponential and Weibull
    [InlineData("=EXPON.DIST(0.2,10,TRUE)", 0.8646647167633873)]
    [InlineData("=EXPON.DIST(0.2,10,FALSE)", 1.353352832366127)]
    [InlineData("=EXPON.DIST(1e-5,1e-5,TRUE)", 9.999999999500002e-11)]
    [InlineData("=EXPON.DIST(-1e-5,10,TRUE)", "#NUM!")]
    [InlineData("=EXPON.DIST(1,0,TRUE)", "#NUM!")]
    [InlineData("=EXPONDIST(1,2,FALSE)", 0.2706705664732254)]
    [InlineData("=WEIBULL.DIST(0.8,1,9,TRUE)", 0.08505277126996899)]
    [InlineData("=WEIBULL.DIST(0.4,1,9,FALSE)", 0.10628097101144768)]
    [InlineData("=WEIBULL.DIST(-1,1,9,FALSE)", "#NUM!")]
    [InlineData("=WEIBULL(0.8,1,9,TRUE)", 0.08505277126996899)]
    // Gamma
    [InlineData("=GAMMA.DIST(10,9,2,TRUE)", 0.06809363472184855)]
    [InlineData("=GAMMA.DIST(10,9,2,FALSE)", 0.032639019674079374)]
    [InlineData("=GAMMA.DIST(1e-5,9,2,TRUE)", 5.38226469068969e-54)]
    [InlineData("=GAMMA.DIST(10,9,1000000,TRUE)", 2.755707120924027e-51)]
    [InlineData("=GAMMA.DIST(0,9,2,FALSE)", 0)]
    [InlineData("=GAMMA.DIST(0,0.5,1,FALSE)", "#NUM!")]
    [InlineData("=GAMMA.DIST(10,0,2,TRUE)", "#NUM!")]
    [InlineData("=GAMMADIST(10,9,2,TRUE)", 0.06809363472184855)]
    [InlineData("=GAMMA.INV(0.068094,9,2)", 10.00001119143718)]
    [InlineData("=GAMMA.INV(1e-7,9,2)", 1.4902187371078093)]
    [InlineData("=GAMMA.INV(0.999999,9,2)", 61.91422671617931)]
    [InlineData("=GAMMA.INV(0.5,1000,2)", 1999.3333728539303)]
    [InlineData("=GAMMA.INV(0.5,0.5,2)", 0.4549364231195729)]
    [InlineData("=GAMMA.INV(0,9,2)", 0)]
    [InlineData("=GAMMA.INV(1,9,2)", "#NUM!")]
    [InlineData("=GAMMAINV(0.5,1,1)", 0.6931471805599453)]
    [InlineData("=GAMMA(2.5)", 1.329340388179137)]
    [InlineData("=GAMMA(-3.75)", 0.26786612886141664)]
    [InlineData("=GAMMA(1e-10)", 9999999999.422787)]
    [InlineData("=GAMMA(-1e-10)", -10000000000.577215)]
    [InlineData("=GAMMA(-18.1111)", -1.0372074185215364e-15)]
    [InlineData("=GAMMA(100)", 9.332621544394421e+155)]
    [InlineData("=GAMMA(0)", "#NUM!")]
    [InlineData("=GAMMA(-2)", "#NUM!")]
    [InlineData("=GAMMA(172)", "#NUM!")]
    [InlineData("=GAMMA(TRUE)", 1)]
    [InlineData("=GAMMALN(4)", 1.791759469228055)]
    [InlineData("=GAMMALN(1.0000001)", -5.772155829918508e-08)]
    [InlineData("=GAMMALN(0.9999999)", 5.7721574684441944e-08)]
    [InlineData("=GAMMALN(10000000000)", 220258509288.81058)]
    [InlineData("=GAMMALN(\" 4 \")", 1.791759469228055)]
    [InlineData("=GAMMALN(0)", "#NUM!")]
    [InlineData("=GAMMALN.PRECISE(-0.5)", "#NUM!")]
    [InlineData("=GAMMALN.PRECISE(23.4)", 49.720154482211285)]
    // Beta
    [InlineData("=BETA.DIST(3,2,3,TRUE,1,10)", 0.21582075903063558)]
    [InlineData("=BETA.DIST(3,2,3,FALSE,1,10)", 0.17924096936442618)]
    [InlineData("=BETA.DIST(0.234,1.7,2.5,TRUE,0.15,1.2)", 0.04773722139554878)]
    [InlineData("=BETA.DIST(33587,3,600,TRUE,2000,5489582)", 0.6734553797493481)]
    [InlineData("=BETA.DIST(0.234001,2,2.5,TRUE,0.234,1.2)", 4.688385940594357e-12)]
    [InlineData("=BETA.DIST(1.199999,2,2.5,TRUE,0.234,1.2)", 0.9999999999999962)]
    [InlineData("=BETA.DIST(1e-12,2,2.5,TRUE,0,1.2)", 3.038194444441917e-24)]
    [InlineData("=BETA.DIST(0.9,0.1,10,TRUE,0,1)", 0.9999999999985647)]
    [InlineData("=BETA.DIST(0.234,0.0001,2,TRUE,0.15,1.2)", 0.9998394357956406)]
    [InlineData("=BETA.DIST(0.234,0.0001,0.0001,TRUE,0.15,1.2)", 0.4998779049300665)]
    [InlineData("=BETA.DIST(0.555,100,1,TRUE,0.1,1)", 2.3817091316180232e-30)]
    [InlineData("=BETA.DIST(0.555,1000,1000,TRUE,0.1,1)", 0.6903551301985376)]
    [InlineData("=BETA.DIST(0.234,2,2.5,TRUE,0.234,1.2)", 0)]
    [InlineData("=BETA.DIST(0.234,2,2.5,TRUE,0.234,0.234)", "#NUM!")]
    [InlineData("=BETA.DIST(0.23,2,2.5,TRUE,0.234,1.2)", "#NUM!")]
    [InlineData("=BETA.DIST(0.234,0,2.5,TRUE,0.15,1.2)", "#NUM!")]
    [InlineData("=BETA.DIST(1,1,1,TRUE)", 1)]
    [InlineData("=BETA.DIST(2,8,10,FALSE,1,3)", 1.4837646484375)]
    [InlineData("=BETADIST(0.234,1.7,2.5,0.15,1.2)", 0.04773722139554878)]
    [InlineData("=BETA.INV(0.21582075903063558,2,3,1,10)", 3)]
    [InlineData("=BETA.INV(0.17924096936442618,2,3,1,10)", 2.79085340092616)]
    [InlineData("=BETA.INV(0.6734553797493481,3,600,2000,5489582)", 33587.000000000226)]
    [InlineData("=BETA.INV(4.688385940594357e-12,2,2.5,0.234,1.2)", 0.234001)]
    [InlineData("=BETA.INV(0.9998394357956406,0.0001,2,0.15,1.2)", 0.23400000000003063)]
    [InlineData("=BETA.INV(2.3817091316180232e-30,100,1,0.1,1)", 0.555)]
    [InlineData("=BETA.INV(0.0001,1,2,0,1)", 5.000125006250391e-05)]
    [InlineData("=BETA.INV(0.99999,1,2,0,1)", 0.9968377223398388)]
    [InlineData("=BETA.INV(0,1,2,0,1)", "#NUM!")]
    [InlineData("=BETA.INV(1,1,2,0,1)", "#NUM!")]
    [InlineData("=BETA.INV(0.5,1,2,1,1)", "#NUM!")]
    [InlineData("=BETAINV(0.04773722139554878,1.7,2.5,0.15,1.2)", 0.234)]
    // Chi-square (degrees of freedom are truncated)
    [InlineData("=CHISQ.DIST(12.3,4,TRUE)", 0.9847456053442304)]
    [InlineData("=CHISQ.DIST(8,3.34,TRUE)", 0.9539882943107687)]
    [InlineData("=CHISQ.DIST(12.3,7.2,FALSE)", 0.030107210682426052)]
    [InlineData("=CHISQ.DIST(-1,4,TRUE)", "#NUM!")]
    [InlineData("=CHISQ.DIST(1,0.5,TRUE)", "#NUM!")]
    [InlineData("=CHISQ.DIST.RT(12.3,4)", 0.015254394655769613)]
    [InlineData("=CHISQ.DIST.RT(12.3,7.2)", 0.09111488600031305)]
    [InlineData("=CHIDIST(12.3,4)", 0.015254394655769613)]
    [InlineData("=CHISQ.INV(0.9847456053442304,4)", 12.299999999999997)]
    [InlineData("=CHISQ.INV(0,4)", 0)]
    [InlineData("=CHISQ.INV(1,4)", "#NUM!")]
    [InlineData("=CHISQ.INV.RT(0.015254394655769613,4)", 12.3)]
    [InlineData("=CHISQ.INV.RT(0,4)", "#NUM!")]
    [InlineData("=CHIINV(0.015254394655769613,4)", 12.3)]
    // F (degrees of freedom are truncated)
    [InlineData("=F.DIST(0.8,2.4,3,TRUE)", 0.4733220523344033)]
    [InlineData("=F.DIST(0.7,4,2.4,FALSE)", 0.40509259259259256)]
    [InlineData("=F.DIST(-0.7,4,2,FALSE)", "#NUM!")]
    [InlineData("=F.DIST(0.7,0,2,FALSE)", "#NUM!")]
    [InlineData("=F.DIST.RT(0.8,2.4,3)", 0.5266779476655967)]
    [InlineData("=FDIST(0.7,4,2.4)", 0.6597222222222223)]
    [InlineData("=F.INV(0.4733220523344033,2.4,3)", 0.8000000000000002)]
    [InlineData("=F.INV(0.40509259259259256,4,2.4)", 0.8753979410454287)]
    [InlineData("=F.INV(0,4,2)", 0)]
    [InlineData("=F.INV.RT(0.6597222222222223,4,2.4)", 0.6999999999999997)]
    [InlineData("=FINV(0.5266779476655967,2,3)", 0.8000000000000002)]
    [InlineData("=FINV(0,2,3)", "#NUM!")]
    // Student's t (degrees of freedom are truncated)
    [InlineData("=T.DIST(60,1,TRUE)", 0.9946953263673767)]
    [InlineData("=T.DIST(8,3,FALSE)", 0.0007369065209469266)]
    [InlineData("=T.DIST(3,5.2,TRUE)", 0.9849503760512687)]
    [InlineData("=T.DIST(-2,5.2,FALSE)", 0.0650903103262165)]
    [InlineData("=T.DIST(1.959999998,60,TRUE)", 0.9726775350120396)]
    [InlineData("=T.DIST(0,10,FALSE)", 0.38910838396603115)]
    [InlineData("=T.DIST(-1000000,10,TRUE)", 1.230468749943611e-56)]
    [InlineData("=T.DIST(1e-6,11,TRUE)", 0.500000389989757)]
    [InlineData("=T.DIST(1,0.999999,TRUE)", "#NUM!")]
    [InlineData("=T.DIST(1,1.000001,TRUE)", 0.75)]
    [InlineData("=T.DIST.2T(1,10)", 0.34089313230205986)]
    [InlineData("=T.DIST.2T(1e-6,10)", 0.9999992217832321)]
    [InlineData("=T.DIST.2T(-1,10)", "#NUM!")]
    [InlineData("=T.DIST.RT(1000000,10)", 1.230468749943611e-56)]
    [InlineData("=T.DIST.RT(-2,5.2)", 0.9490302605850708)]
    [InlineData("=TDIST(1,10,2)", 0.34089313230205986)]
    [InlineData("=TDIST(1,10,1)", 0.17044656615102993)]
    [InlineData("=TDIST(-1,10,1)", "#NUM!")]
    [InlineData("=TDIST(1,10,3)", "#NUM!")]
    [InlineData("=T.INV(0.9946953263673767,1)", 59.99999999999939)]
    [InlineData("=T.INV(0.0007369065209469266,3)", -11.332581353101341)]
    [InlineData("=T.INV(0.500000389989757,11)", 9.999999999666944e-07)]
    [InlineData("=T.INV(1.230468749943611e-56,10)", -1000000.0000000001)]
    [InlineData("=T.INV(1e-7,10)", -12.492209687558763)]
    [InlineData("=T.INV(0.9999999,10)", 12.492209688254745)]
    [InlineData("=T.INV(0.5,1000000)", 0)]
    [InlineData("=T.INV(1,10)", "#NUM!")]
    [InlineData("=T.INV.2T(0.05,10)", 2.2281388519862744)]
    [InlineData("=T.INV.2T(0.025,1)", 25.45169957935708)]
    [InlineData("=T.INV.2T(0.9999999,10)", 1.2849890167888884e-07)]
    [InlineData("=T.INV.2T(0.5,1000)", 0.674735164607012)]
    [InlineData("=T.INV.2T(1,10)", 0)]
    [InlineData("=T.INV.2T(0,10)", "#NUM!")]
    [InlineData("=TINV(0.05,10)", 2.2281388519862744)]
    // Fisher transformation
    [InlineData("=FISHER(0.75)", 0.9729550745276566)]
    [InlineData("=FISHER(1e-10)", 1.000000082640371e-10)]
    [InlineData("=FISHER(1)", "#NUM!")]
    [InlineData("=FISHER(TRUE)", "#NUM!")]
    [InlineData("=FISHERINV(1e-7)", 1.0000000005039292e-07)]
    [InlineData("=FISHERINV(-10)", -0.9999999958776927)]
    [InlineData("=FISHERINV(1000)", 1)]
    [InlineData("=FISHERINV(\"0.5\")", 0.46211715726000974)]
    // Error function: logical values are not numbers here
    [InlineData("=ERF(1)", 0.8427007929497149)]
    [InlineData("=ERF(0.1)", 0.11246291601828493)]
    [InlineData("=ERF(-5)", -0.9999999999984626)]
    [InlineData("=ERF(0,0.5)", 0.5204998778130465)]
    [InlineData("=ERF(0.3333,0.7777)", 0.3659810414127489)]
    [InlineData("=ERF(6,-3)", -1.9999779095030012)]
    [InlineData("=ERF(\"1\",\"2\")", 0.15262147206923782)]
    [InlineData("=ERF(TRUE)", "#VALUE!")]
    [InlineData("=ERF(1,TRUE)", "#VALUE!")]
    [InlineData("=ERF(\"str\")", "#VALUE!")]
    [InlineData("=ERFC(5)", 1.537459794428034e-12)]
    [InlineData("=ERFC(6)", 2.1519736712498925e-17)]
    [InlineData("=ERFC(7)", 4.1838256077794166e-23)]
    [InlineData("=ERFC(-1)", 1.8427007929497148)]
    [InlineData("=ERFC(0.1)", 0.887537083981715)]
    [InlineData("=ERFC.PRECISE(4.5)", 1.9661604415428865e-10)]
    [InlineData("=ERF.PRECISE(-0.1)", -0.11246291601828493)]
    [InlineData("=ERFC(FALSE)", "#VALUE!")]
    [InlineData("=ERF(1/0)", "#DIV/0!")]
    public void MatchesExcel(string formula, object expected)
    {
        AssertExcel(expected, _wb.Evaluate(formula));
    }

    [Fact]
    public void The_first_bad_argument_decides_the_error()
    {
        AssertExcel("#DIV/0!", _wb.Evaluate("=NORM.DIST(1/0,0,-1,TRUE)"));
        AssertExcel("#N/A", _wb.Evaluate("=BETA.DIST(3,#N/A,3,#REF!,1,10)"));
        AssertExcel("#VALUE!", _wb.Evaluate("=GAMMA.DIST(10,9,\"x\",1/0)"));
    }

    [Fact]
    public void An_array_argument_gives_an_array_of_results()
    {
        var result = _wb.Evaluate("=NORM.S.DIST({0,1},TRUE)");
        Assert.Equal(CellValueKind.Array, result.Kind);
        var values = result.AsArray();
        AssertExcel(0.5, values[0, 0]);
        AssertExcel(0.841344746068543, values[0, 1]);
    }
}
