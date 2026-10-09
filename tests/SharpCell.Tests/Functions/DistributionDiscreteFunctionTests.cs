using SharpCell;

namespace SharpCell.Tests.Functions;

/// <summary>Discrete distributions and permutations; expected values are Excel's.</summary>
public class DistributionDiscreteFunctionTests
{
    private readonly Workbook _wb = new();

    public DistributionDiscreteFunctionTests()
    {
        _wb.AddSheet("S");
    }

    [Theory]
    // Binomial (counts truncate)
    [InlineData("=BINOM.DIST(7,8,0.7,TRUE)", 0.94235199)]
    [InlineData("=BINOM.DIST(3,17,0.2,FALSE)", 0.23925373020405763)]
    [InlineData("=BINOM.DIST(1.2,2.7,0.1,TRUE)", 0.99)]
    [InlineData("=BINOM.DIST(2.99,10,0.5,TRUE)", 0.0546875)]
    [InlineData("=BINOM.DIST(1,3000,0.0001,TRUE)", 0.9630714665120287)]
    [InlineData("=BINOM.DIST(1000,10000,0.1,TRUE)", 0.508421039265265)]
    [InlineData("=BINOM.DIST(9999,10000,0.999,TRUE)", 0.9999548266540229)]
    [InlineData("=BINOM.DIST(1,1000000,0.00001,TRUE)", 0.0004993787976951216)]
    [InlineData("=BINOM.DIST(2,10,0.000001,TRUE)", 0.9999999999999999)]
    [InlineData("=BINOM.DIST(2,10,0.999999,TRUE)", 4.499992001038836e-47)]
    [InlineData("=BINOM.DIST(1,100,0.99,TRUE)", 9.90100000000073e-197)]
    [InlineData("=BINOM.DIST(0,0,0,TRUE)", 1)]
    [InlineData("=BINOM.DIST(2,10,1,TRUE)", 0)]
    [InlineData("=BINOM.DIST(5,10,0.3,FALSE)", 0.10291934520000003)]
    [InlineData("=BINOM.DIST(,10,0.3,TRUE)", 0.028247524899999994)]
    [InlineData("=BINOM.DIST(,,,)", 1)]
    [InlineData("=BINOM.DIST(10,2,0.3,TRUE)", "#NUM!")]
    [InlineData("=BINOM.DIST(-2,10,0.3,TRUE)", "#NUM!")]
    [InlineData("=BINOM.DIST(2,10,2,TRUE)", "#NUM!")]
    [InlineData("=BINOM.DIST(10000000000,1E+30,0.3,TRUE)", "#NUM!")]
    [InlineData("=BINOM.DIST(7,8,0.7,\"abc\")", "#VALUE!")]
    [InlineData("=BINOMDIST(7,8,0.7,FALSE)", 0.19765031999999993)]
    [InlineData("=BINOM.DIST.RANGE(8,0.7,7,8)", 0.2552983299999998)]
    [InlineData("=BINOM.DIST.RANGE(10.7,0.1,1.2,5.5)", 0.6511746573)]
    [InlineData("=BINOM.DIST.RANGE(3000,0.0001,1,1)", 0.2222643587611419)]
    [InlineData("=BINOM.DIST.RANGE(10,0.000001,2,3)", 4.499976000042004e-11)]
    [InlineData("=BINOM.DIST.RANGE(10,0.999999,2,3)", 1.1999968502442456e-40)]
    [InlineData("=BINOM.DIST.RANGE(2000,0.1,1995,1996)", 0)]
    [InlineData("=BINOM.DIST.RANGE(10,0.3,5)", 0.10291934520000003)]
    [InlineData("=BINOM.DIST.RANGE(10,0.3,0,10)", 1)]
    [InlineData("=BINOM.DIST.RANGE(10,0.3,5,3)", "#NUM!")]
    [InlineData("=BINOM.DIST.RANGE(10,0.3,5,12)", "#NUM!")]
    [InlineData("=BINOM.DIST.RANGE(2,0.1,2,)", "#NUM!")]
    [InlineData("=BINOM.DIST.RANGE(100000,0.4,30000,50000)", 1)]
    [InlineData("=BINOM.INV(10,0.7,0.8)", 8)]
    [InlineData("=BINOM.INV(8,0.7,0.999999)", 8)]
    [InlineData("=BINOM.INV(1,0.7,0.8)", 1)]
    [InlineData("=BINOM.INV(100000,0.7,0.8)", 70122)]
    [InlineData("=BINOM.INV(100000,0.1,0.5)", 10000)]
    [InlineData("=BINOM.INV(10,0.999999,0.8)", 10)]
    [InlineData("=BINOM.INV(10.9999,0.7,0.8)", 8)]
    [InlineData("=BINOM.INV(10,0.5,0.1)", 3)]
    [InlineData("=BINOM.INV(10,0.7,0)", "#NUM!")]
    [InlineData("=BINOM.INV(10,0.7,1)", "#NUM!")]
    [InlineData("=BINOM.INV(10,1,0.8)", "#NUM!")]
    [InlineData("=BINOM.INV(-10,0.7,0.8)", "#NUM!")]
    [InlineData("=CRITBINOM(10,0.5,0.25)", 4)]
    // Negative binomial
    [InlineData("=NEGBINOM.DIST(7,8,0.7,TRUE)", 0.949987459946224)]
    [InlineData("=NEGBINOM.DIST(3,17,0.2,FALSE)", 6.50284892159999e-10)]
    [InlineData("=NEGBINOM.DIST(1.2,2.7,0.1,TRUE)", 0.02800000000000001)]
    [InlineData("=NEGBINOM.DIST(0,100,0.01,TRUE)", 1.0000000000000348e-200)]
    [InlineData("=NEGBINOM.DIST(99,100,0.01,TRUE)", 1.690659363904351e-142)]
    [InlineData("=NEGBINOM.DIST(2,10,0.999999,TRUE)", 0.9999999999999998)]
    [InlineData("=NEGBINOM.DIST(2,10000000000,0.3,TRUE)", 0)]
    [InlineData("=NEGBINOM.DIST(1000000,10,0.5,TRUE)", 1)]
    [InlineData("=NEGBINOM.DIST(10,2,0.3,TRUE)", 0.914974950051)]
    [InlineData("=NEGBINOM.DIST(5,0,0.5,TRUE)", "#NUM!")]
    [InlineData("=NEGBINOM.DIST(2,10,0,TRUE)", "#NUM!")]
    [InlineData("=NEGBINOM.DIST(2,10,1,TRUE)", "#NUM!")]
    [InlineData("=NEGBINOM.DIST(-2,10,0.3,TRUE)", "#NUM!")]
    [InlineData("=NEGBINOMDIST(5,10,0.3)", 0.001986857959086001)]
    // Hypergeometric
    [InlineData("=HYPGEOM.DIST(1.8,4.2,8.2,22,TRUE)", 0.5349282296650714)]
    [InlineData("=HYPGEOM.DIST(1.8,4.2,8.2,18.34,FALSE)", 0.31372549019607815)]
    [InlineData("=HYPGEOM.DIST(1,4,8,20,TRUE)", 0.4654282765737875)]
    [InlineData("=HYPGEOM.DIST(3,4,8,20,TRUE)", 0.9855521155830753)]
    [InlineData("=HYPGEOM.DIST(5,4,8,20,TRUE)", "#NUM!")]
    [InlineData("=HYPGEOM.DIST(1,25,8,20,TRUE)", "#NUM!")]
    [InlineData("=HYPGEOMDIST(1,4,8,20)", 0.3632610939112487)]
    // Poisson
    [InlineData("=POISSON.DIST(2,5,TRUE)", 0.12465201948308113)]
    [InlineData("=POISSON.DIST(2,5,FALSE)", 0.08422433748856833)]
    [InlineData("=POISSON.DIST(0,0,FALSE)", 1)]
    [InlineData("=POISSON.DIST(1,0,FALSE)", 0)]
    [InlineData("=POISSON.DIST(0,0.001,TRUE)", 0.999000499833375)]
    [InlineData("=POISSON.DIST(100,100,TRUE)", 0.5265621985299984)]
    [InlineData("=POISSON.DIST(1000000,1000000,TRUE)", 0.5002659614862837)]
    [InlineData("=POISSON.DIST(10,1000000,TRUE)", 0)]
    [InlineData("=POISSON.DIST(2.999999,5,TRUE)", 0.12465201948308113)]
    [InlineData("=POISSON.DIST(-0.000001,5,TRUE)", "#NUM!")]
    [InlineData("=POISSON.DIST(3,-0.000001,TRUE)", "#NUM!")]
    [InlineData("=POISSON(2,5,TRUE)", 0.12465201948308113)]
    // Permutations
    [InlineData("=PERMUT(12,6)", 665280)]
    [InlineData("=PERMUT(12.7,4.6)", 11880)]
    [InlineData("=PERMUT(3,0)", 1)]
    [InlineData("=PERMUT(1,2)", "#NUM!")]
    [InlineData("=PERMUT(0,0)", "#NUM!")]
    [InlineData("=PERMUT(2,-1)", "#NUM!")]
    [InlineData("=PERMUT(1000,1000)", "#NUM!")]
    [InlineData("=PERMUTATIONA(12,6)", 2985984)]
    [InlineData("=PERMUTATIONA(0,2)", 0)]
    [InlineData("=PERMUTATIONA(3,0)", 1)]
    [InlineData("=PERMUTATIONA(-1,2)", "#NUM!")]
    public void MatchesExcel(string formula, object expected)
    {
        DistributionFunctionTests.AssertExcel(expected, _wb.Evaluate(formula));
    }
}
