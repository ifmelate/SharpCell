using System;
using SharpCell.Functions;

namespace SharpCell.Tests.Functions;

public class SpecialTests
{
    // Reference values computed with mpmath at 60 digits.
    [Theory]
    [InlineData("Erf", -30.0, 0.0, 0.0, -1.0)]
    [InlineData("Erfc", -30.0, 0.0, 0.0, 2.0)]
    [InlineData("Erf", -6.0, 0.0, 0.0, -1.0)]
    [InlineData("Erfc", -6.0, 0.0, 0.0, 2.0)]
    [InlineData("Erf", -3.5, 0.0, 0.0, -0.9999992569016276)]
    [InlineData("Erfc", -3.5, 0.0, 0.0, 1.9999992569016276)]
    [InlineData("Erf", -1.2, 0.0, 0.0, -0.9103139782296353)]
    [InlineData("Erfc", -1.2, 0.0, 0.0, 1.9103139782296354)]
    [InlineData("Erf", -1.0, 0.0, 0.0, -0.8427007929497149)]
    [InlineData("Erfc", -1.0, 0.0, 0.0, 1.8427007929497148)]
    [InlineData("Erf", -0.999, 0.0, 0.0, -0.8422852702064969)]
    [InlineData("Erfc", -0.999, 0.0, 0.0, 1.842285270206497)]
    [InlineData("Erf", -0.5, 0.0, 0.0, -0.5204998778130465)]
    [InlineData("Erfc", -0.5, 0.0, 0.0, 1.5204998778130465)]
    [InlineData("Erf", -0.1, 0.0, 0.0, -0.1124629160182849)]
    [InlineData("Erfc", -0.1, 0.0, 0.0, 1.1124629160182848)]
    [InlineData("Erf", -1e-10, 0.0, 0.0, -1.1283791670955126e-10)]
    [InlineData("Erfc", -1e-10, 0.0, 0.0, 1.000000000112838)]
    [InlineData("Erf", 0.0, 0.0, 0.0, 0.0)]
    [InlineData("Erfc", 0.0, 0.0, 0.0, 1.0)]
    [InlineData("Erf", 1e-300, 0.0, 0.0, 1.1283791670955126e-300)]
    [InlineData("Erfc", 1e-300, 0.0, 0.0, 1.0)]
    [InlineData("Erf", 1e-08, 0.0, 0.0, 1.1283791670955126e-08)]
    [InlineData("Erfc", 1e-08, 0.0, 0.0, 0.9999999887162083)]
    [InlineData("Erf", 0.1, 0.0, 0.0, 0.1124629160182849)]
    [InlineData("Erfc", 0.1, 0.0, 0.0, 0.887537083981715)]
    [InlineData("Erf", 0.745, 0.0, 0.0, 0.7079289200957377)]
    [InlineData("Erfc", 0.745, 0.0, 0.0, 0.29207107990426234)]
    [InlineData("Erf", 0.999, 0.0, 0.0, 0.8422852702064969)]
    [InlineData("Erfc", 0.999, 0.0, 0.0, 0.15771472979350307)]
    [InlineData("Erf", 1.0, 0.0, 0.0, 0.8427007929497149)]
    [InlineData("Erfc", 1.0, 0.0, 0.0, 0.15729920705028513)]
    [InlineData("Erf", 1.0001, 0.0, 0.0, 0.8427422995485203)]
    [InlineData("Erfc", 1.0001, 0.0, 0.0, 0.1572577004514797)]
    [InlineData("Erf", 1.5, 0.0, 0.0, 0.9661051464753108)]
    [InlineData("Erfc", 1.5, 0.0, 0.0, 0.033894853524689274)]
    [InlineData("Erf", 2.0, 0.0, 0.0, 0.9953222650189527)]
    [InlineData("Erfc", 2.0, 0.0, 0.0, 0.004677734981047266)]
    [InlineData("Erf", 2.5, 0.0, 0.0, 0.999593047982555)]
    [InlineData("Erfc", 2.5, 0.0, 0.0, 0.0004069520174449589)]
    [InlineData("Erf", 3.0, 0.0, 0.0, 0.9999779095030014)]
    [InlineData("Erfc", 3.0, 0.0, 0.0, 2.209049699858544e-05)]
    [InlineData("Erf", 4.0, 0.0, 0.0, 0.9999999845827421)]
    [InlineData("Erfc", 4.0, 0.0, 0.0, 1.541725790028002e-08)]
    [InlineData("Erf", 5.0, 0.0, 0.0, 0.9999999999984626)]
    [InlineData("Erfc", 5.0, 0.0, 0.0, 1.537459794428035e-12)]
    [InlineData("Erf", 6.0, 0.0, 0.0, 1.0)]
    [InlineData("Erfc", 6.0, 0.0, 0.0, 2.1519736712498913e-17)]
    [InlineData("Erf", 10.0, 0.0, 0.0, 1.0)]
    [InlineData("Erfc", 10.0, 0.0, 0.0, 2.088487583762545e-45)]
    [InlineData("Erf", 20.0, 0.0, 0.0, 1.0)]
    [InlineData("Erfc", 20.0, 0.0, 0.0, 5.395865611607901e-176)]
    [InlineData("Erf", 26.0, 0.0, 0.0, 1.0)]
    [InlineData("Erfc", 26.0, 0.0, 0.0, 5.663192408856143e-296)]
    [InlineData("NormalCdf", -38.0, 0.0, 0.0, 2.88542835e-316)]
    [InlineData("NormalCdf", -37.0, 0.0, 0.0, 5.725571222524577e-300)]
    [InlineData("NormalCdf", -20.0, 0.0, 0.0, 2.7536241186062337e-89)]
    [InlineData("NormalCdf", -10.0, 0.0, 0.0, 7.619853024160525e-24)]
    [InlineData("NormalCdf", -5.0, 0.0, 0.0, 2.866515718791939e-07)]
    [InlineData("NormalCdf", -3.0, 0.0, 0.0, 0.0013498980316300946)]
    [InlineData("NormalCdf", -1.5, 0.0, 0.0, 0.06680720126885807)]
    [InlineData("NormalCdf", -1.41, 0.0, 0.0, 0.0792698414533924)]
    [InlineData("NormalCdf", -1.4, 0.0, 0.0, 0.08075665923377107)]
    [InlineData("NormalCdf", -1.0, 0.0, 0.0, 0.15865525393145705)]
    [InlineData("NormalCdf", -1e-06, 0.0, 0.0, 0.4999996010577196)]
    [InlineData("NormalCdf", 0.0, 0.0, 0.0, 0.5)]
    [InlineData("NormalCdf", 1e-07, 0.0, 0.0, 0.500000039894228)]
    [InlineData("NormalCdf", 0.5, 0.0, 0.0, 0.6914624612740131)]
    [InlineData("NormalCdf", 1.39, 0.0, 0.0, 0.917735561322331)]
    [InlineData("NormalCdf", 1.41, 0.0, 0.0, 0.9207301585466076)]
    [InlineData("NormalCdf", 2.0, 0.0, 0.0, 0.9772498680518208)]
    [InlineData("NormalCdf", 3.0, 0.0, 0.0, 0.9986501019683699)]
    [InlineData("NormalCdf", 5.0, 0.0, 0.0, 0.9999997133484281)]
    [InlineData("NormalCdf", 8.2, 0.0, 0.0, 0.9999999999999999)]
    [InlineData("NormalCdf", 9.0, 0.0, 0.0, 1.0)]
    [InlineData("NormalQuantile", 1e-300, 0.0, 0.0, -37.0470962993612)]
    [InlineData("NormalQuantile", 1e-100, 0.0, 0.0, -21.273453560965326)]
    [InlineData("NormalQuantile", 1e-20, 0.0, 0.0, -9.262340089798407)]
    [InlineData("NormalQuantile", 7.619853024160475e-24, 0.0, 0.0, -10.0)]
    [InlineData("NormalQuantile", 1e-07, 0.0, 0.0, -5.1993375821928165)]
    [InlineData("NormalQuantile", 0.001, 0.0, 0.0, -3.0902323061678136)]
    [InlineData("NormalQuantile", 0.02, 0.0, 0.0, -2.053748910631823)]
    [InlineData("NormalQuantile", 0.02425, 0.0, 0.0, -1.972961051311885)]
    [InlineData("NormalQuantile", 0.025, 0.0, 0.0, -1.9599639845400543)]
    [InlineData("NormalQuantile", 0.1, 0.0, 0.0, -1.2815515655446004)]
    [InlineData("NormalQuantile", 0.3, 0.0, 0.0, -0.5244005127080408)]
    [InlineData("NormalQuantile", 0.4999999, 0.0, 0.0, -2.5066282747031063e-07)]
    [InlineData("NormalQuantile", 0.5, 0.0, 0.0, 0.0)]
    [InlineData("NormalQuantile", 0.5000003989422803, 0.0, 0.0, 9.999999998007829e-07)]
    [InlineData("NormalQuantile", 0.75, 0.0, 0.0, 0.6744897501960817)]
    [InlineData("NormalQuantile", 0.9, 0.0, 0.0, 1.2815515655446006)]
    [InlineData("NormalQuantile", 0.975, 0.0, 0.0, 1.9599639845400538)]
    [InlineData("NormalQuantile", 0.97575, 0.0, 0.0, 1.972961051311885)]
    [InlineData("NormalQuantile", 0.999999, 0.0, 0.0, 4.753424308817087)]
    [InlineData("NormalQuantile", 0.9999999999999999, 0.0, 0.0, 8.209536151601387)]
    [InlineData("LogGamma", 1e-300, 0.0, 0.0, 690.7755278982137)]
    [InlineData("LogGamma", 1e-10, 0.0, 0.0, 23.025850929882736)]
    [InlineData("LogGamma", 1e-07, 0.0, 0.0, 16.118095593236763)]
    [InlineData("LogGamma", 0.001, 0.0, 0.0, 6.907178885383853)]
    [InlineData("LogGamma", 0.1, 0.0, 0.0, 2.252712651734206)]
    [InlineData("LogGamma", 0.4999, 0.0, 0.0, 0.5725613186041184)]
    [InlineData("LogGamma", 0.5, 0.0, 0.0, 0.5723649429247001)]
    [InlineData("LogGamma", 0.9999999, 0.0, 0.0, 5.772157468444193e-08)]
    [InlineData("LogGamma", 1.0000001, 0.0, 0.0, -5.772155829918507e-08)]
    [InlineData("LogGamma", 1.2, 0.0, 0.0, -0.08537409000331583)]
    [InlineData("LogGamma", 1.5, 0.0, 0.0, -0.12078223763524522)]
    [InlineData("LogGamma", 1.7, 0.0, 0.0, -0.09580769740706588)]
    [InlineData("LogGamma", 2.22, 0.0, 0.0, 0.107947496900865)]
    [InlineData("LogGamma", 2.5, 0.0, 0.0, 0.2846828704729192)]
    [InlineData("LogGamma", 2.50001, 0.0, 0.0, 0.2846899020638435)]
    [InlineData("LogGamma", 3.9999999, 0.0, 0.0, 1.7917593436162897)]
    [InlineData("LogGamma", 4.0000001, 0.0, 0.0, 1.7917595948398237)]
    [InlineData("LogGamma", 7.5, 0.0, 0.0, 7.534364236758733)]
    [InlineData("LogGamma", 9.9999999, 0.0, 0.0, 12.801827254906213)]
    [InlineData("LogGamma", 10.0000001, 0.0, 0.0, 12.801827705256727)]
    [InlineData("LogGamma", 14.9, 0.0, 0.0, 24.924132002217277)]
    [InlineData("LogGamma", 15.0, 0.0, 0.0, 25.19122118273868)]
    [InlineData("LogGamma", 15.1, 0.0, 0.0, 25.458999750992664)]
    [InlineData("LogGamma", 23.4, 0.0, 0.0, 49.72015448221128)]
    [InlineData("LogGamma", 100.0, 0.0, 0.0, 359.1342053695754)]
    [InlineData("LogGamma", 1000.0, 0.0, 0.0, 5905.220423209181)]
    [InlineData("LogGamma", 45675.0, 0.0, 0.0, 444381.6232505876)]
    [InlineData("LogGamma", 10000000000.0, 0.0, 0.0, 220258509288.81058)]
    [InlineData("LogGamma", 1e+300, 0.0, 0.0, 6.897755278982137e+302)]
    [InlineData("Gamma", -170.5, 0.0, 0.0, -3.3127395215386074e-308)]
    [InlineData("Gamma", -18.1111, 0.0, 0.0, -1.0372074185215358e-15)]
    [InlineData("Gamma", -3.75, 0.0, 0.0, 0.2678661288614166)]
    [InlineData("Gamma", -3.2, 0.0, 0.0, 0.689056412005979)]
    [InlineData("Gamma", -1.5, 0.0, 0.0, 2.363271801207355)]
    [InlineData("Gamma", -0.5, 0.0, 0.0, -3.544907701811032)]
    [InlineData("Gamma", -1e-10, 0.0, 0.0, -10000000000.577215)]
    [InlineData("Gamma", 1e-10, 0.0, 0.0, 9999999999.422785)]
    [InlineData("Gamma", 0.3, 0.0, 0.0, 2.991568987687591)]
    [InlineData("Gamma", 0.5, 0.0, 0.0, 1.772453850905516)]
    [InlineData("Gamma", 1.5, 0.0, 0.0, 0.886226925452758)]
    [InlineData("Gamma", 2.5, 0.0, 0.0, 1.329340388179137)]
    [InlineData("Gamma", 5.0, 0.0, 0.0, 24.0)]
    [InlineData("Gamma", 10.0, 0.0, 0.0, 362880.0)]
    [InlineData("Gamma", 14.5, 0.0, 0.0, 23092317922.31424)]
    [InlineData("Gamma", 15.5, 0.0, 0.0, 334838609873.55646)]
    [InlineData("Gamma", 23.4, 0.0, 0.0, 3.919121530539987e+21)]
    [InlineData("Gamma", 100.0, 0.0, 0.0, 9.332621544394415e+155)]
    [InlineData("Gamma", 140.3, 0.0, 0.0, 4.23150139454543e+239)]
    [InlineData("Gamma", 170.0, 0.0, 0.0, 4.269068009004705e+304)]
    [InlineData("Gamma", 171.5, 0.0, 0.0, 9.4833675668248e+307)]
    [InlineData("LogBeta", 0.5, 0.5, 0.0, 1.1447298858494002)]
    [InlineData("LogBeta", 1.0, 1.0, 0.0, 0.0)]
    [InlineData("LogBeta", 2.0, 3.0, 0.0, -2.4849066497880004)]
    [InlineData("LogBeta", 0.0001, 2.0, 0.0, 9.210240376975849)]
    [InlineData("LogBeta", 0.0001, 0.0001, 0.0, 9.903487536089191)]
    [InlineData("LogBeta", 5.0, 50.0, 0.0, -16.57631448650235)]
    [InlineData("LogBeta", 10.0, 10.0, 0.0, -13.736229227036555)]
    [InlineData("LogBeta", 10.0, 3.0, 0.0, -6.492239835020471)]
    [InlineData("LogBeta", 30.5, 40.25, 0.0, -48.87047935065409)]
    [InlineData("LogBeta", 1000.0, 1000.0, 0.0, -1388.4826016359023)]
    [InlineData("LogBeta", 0.5, 1000000.0, 0.0, -6.335390211057437)]
    [InlineData("LogBeta", 100000.0, 2.5, 0.0, -28.497649541827652)]
    [InlineData("LogBeta", 10000000.0, 10000000.0, 0.0, -13862950.404734597)]
    [InlineData("GammaP", 0.5, 0.01, 0.0, 0.1124629160182849)]
    [InlineData("GammaQ", 0.5, 0.01, 0.0, 0.8875370839817152)]
    [InlineData("GammaP", 0.5, 2.0, 0.0, 0.9544997361036416)]
    [InlineData("GammaQ", 0.5, 2.0, 0.0, 0.04550026389635842)]
    [InlineData("GammaP", 0.5, 50.0, 0.0, 1.0)]
    [InlineData("GammaQ", 0.5, 50.0, 0.0, 1.523970604832105e-23)]
    [InlineData("GammaP", 1.0, 1.0, 0.0, 0.6321205588285577)]
    [InlineData("GammaQ", 1.0, 1.0, 0.0, 0.36787944117144233)]
    [InlineData("GammaP", 2.5, 0.1, 0.0, 0.0008861387888124426)]
    [InlineData("GammaQ", 2.5, 0.1, 0.0, 0.9991138612111875)]
    [InlineData("GammaP", 3.0, 5.0, 0.0, 0.8753479805169189)]
    [InlineData("GammaQ", 3.0, 5.0, 0.0, 0.12465201948308115)]
    [InlineData("GammaP", 9.0, 5.0, 0.0, 0.06809363472184855)]
    [InlineData("GammaQ", 9.0, 5.0, 0.0, 0.9319063652781514)]
    [InlineData("GammaP", 9.0, 5e-06, 0.0, 5.382264690689695e-54)]
    [InlineData("GammaQ", 9.0, 5e-06, 0.0, 1.0)]
    [InlineData("GammaP", 9.0, 20.0, 0.0, 0.997912740950865)]
    [InlineData("GammaQ", 9.0, 20.0, 0.0, 0.0020872590491350187)]
    [InlineData("GammaP", 11.0, 5.0, 0.0, 0.013695268598382939)]
    [InlineData("GammaQ", 11.0, 5.0, 0.0, 0.986304731401617)]
    [InlineData("GammaP", 101.0, 100.0, 0.0, 0.47343780147000153)]
    [InlineData("GammaQ", 101.0, 100.0, 0.0, 0.5265621985299984)]
    [InlineData("GammaP", 1001.0, 1000.0, 0.0, 0.491590632831494)]
    [InlineData("GammaQ", 1001.0, 1000.0, 0.0, 0.508409367168506)]
    [InlineData("GammaP", 0.0001, 0.5, 0.0, 0.9999440197070426)]
    [InlineData("GammaQ", 0.0001, 0.5, 0.0, 5.5980292957401714e-05)]
    [InlineData("GammaP", 1000001.0, 1000000.0, 0.0, 0.49973403851371634)]
    [InlineData("GammaQ", 1000001.0, 1000000.0, 0.0, 0.5002659614862837)]
    [InlineData("GammaP", 1000000.0, 1003000.0, 0.0, 0.9986382593537824)]
    [InlineData("GammaQ", 1000000.0, 1003000.0, 0.0, 0.0013617406462175915)]
    [InlineData("GammaP", 50.0, 400.0, 0.0, 1.0)]
    [InlineData("GammaQ", 50.0, 400.0, 0.0, 1.1366407840501794e-109)]
    [InlineData("GammaP", 4.5, 7.3, 0.0, 0.8974743197859524)]
    [InlineData("GammaQ", 4.5, 7.3, 0.0, 0.10252568021404766)]
    [InlineData("BetaI", 0.5, 0.5, 0.3, 0.36901011956554536)]
    [InlineData("BetaIUpper", 0.5, 0.5, 0.3, 0.6309898804344546)]
    [InlineData("BetaI", 2.0, 3.0, 0.2, 0.18080000000000002)]
    [InlineData("BetaIUpper", 2.0, 3.0, 0.2, 0.8191999999999999)]
    [InlineData("BetaI", 2.0, 3.0, 0.9, 0.9963)]
    [InlineData("BetaIUpper", 2.0, 3.0, 0.9, 0.0036999999999999976)]
    [InlineData("BetaI", 0.0001, 2.0, 0.07, 0.9998270846236621)]
    [InlineData("BetaIUpper", 0.0001, 2.0, 0.07, 0.00017291537633789273)]
    [InlineData("BetaI", 2.0, 0.0001, 0.07, 2.570937503206065e-07)]
    [InlineData("BetaIUpper", 2.0, 0.0001, 0.07, 0.9999997429062497)]
    [InlineData("BetaI", 0.0001, 0.0001, 0.07, 0.499870689741388)]
    [InlineData("BetaIUpper", 0.0001, 0.0001, 0.07, 0.5001293102586121)]
    [InlineData("BetaI", 5.0, 0.5, 1e-11, 2.4609375000102533e-56)]
    [InlineData("BetaIUpper", 5.0, 0.5, 1e-11, 1.0)]
    [InlineData("BetaI", 5.0, 0.5, 0.999, 0.9222819921009667)]
    [InlineData("BetaIUpper", 5.0, 0.5, 0.999, 0.07771800789903324)]
    [InlineData("BetaI", 0.5, 5.0, 0.3, 0.9347377538310918)]
    [InlineData("BetaIUpper", 0.5, 5.0, 0.3, 0.06526224616890818)]
    [InlineData("BetaI", 1000.0, 1000.0, 0.52, 0.9632205167213604)]
    [InlineData("BetaIUpper", 1000.0, 1000.0, 0.52, 0.03677948327863957)]
    [InlineData("BetaI", 1001.0, 9000.0, 0.9, 1.0)]
    [InlineData("BetaIUpper", 1001.0, 9000.0, 0.9, 0.0)]
    [InlineData("BetaIUpper", 999999.0, 2.0, 0.99999, 0.9995006212022842)]
    [InlineData("BetaI", 3.0, 10.0, 0.123, 0.1758341902341374)]
    [InlineData("BetaIUpper", 3.0, 10.0, 0.123, 0.8241658097658626)]
    [InlineData("BetaI", 6.0, 1000.0, 0.0222, 0.999989744603517)]
    [InlineData("BetaIUpper", 6.0, 1000.0, 0.0222, 1.0255396483063521e-05)]
    [InlineData("BetaI", 1000.0, 2.0, 0.5, 4.6756507287011266e-299)]
    [InlineData("BetaIUpper", 1000.0, 2.0, 0.5, 1.0)]
    [InlineData("BetaI", 50.0, 50.0, 0.01, 3.1130219337164254e-72)]
    [InlineData("BetaIUpper", 50.0, 50.0, 0.01, 1.0)]
    [InlineData("BetaI", 100.0, 1.0, 0.5, 7.888609052210118e-31)]
    [InlineData("BetaIUpper", 100.0, 1.0, 0.5, 1.0)]
    public void MatchesHighPrecisionReference(string function, double a, double b, double c, double expected)
    {
        var actual = function switch
        {
            "Erf" => Special.Erf(a),
            "Erfc" => Special.Erfc(a),
            "NormalCdf" => Special.NormalCdf(a),
            "NormalQuantile" => Special.NormalQuantile(a),
            "LogGamma" => Special.LogGamma(a),
            "Gamma" => Special.Gamma(a),
            "LogBeta" => Special.LogBeta(a, b),
            "GammaP" => Special.GammaP(a, b),
            "GammaQ" => Special.GammaQ(a, b),
            "BetaI" => Special.BetaRegularized(a, b, c, 1 - c),
            "BetaIUpper" => Special.BetaRegularized(a, b, c, 1 - c, upper: true),
            _ => throw new ArgumentException(function),
        };

        AssertClose(expected, actual, 1e-13);
    }

    [Theory]
    [InlineData(0.5, 1e-10)]
    [InlineData(0.5, 0.3)]
    [InlineData(0.5, 0.999)]
    [InlineData(1, 0.6321205588285577)]
    [InlineData(9, 0.06809363472184855)]
    [InlineData(9, 1e-7)]
    [InlineData(9, 0.999999)]
    [InlineData(1000, 0.5)]
    [InlineData(0.01, 0.4)]
    [InlineData(2.5, 1e-300)]
    [InlineData(1e6, 0.25)]
    public void InverseGammaPInvertsGammaP(double a, double p)
    {
        var x = Special.InverseGammaP(a, p, 1 - p);
        AssertClose(p, Special.GammaP(a, x), 1e-13);
        var upper = Special.InverseGammaP(a, 1 - p, p);
        AssertClose(p, Special.GammaQ(a, upper), 1e-13);
    }

    [Theory]
    [InlineData(2, 3, 0.3)]
    [InlineData(2, 3, 1e-12)]
    [InlineData(0.5, 2, 0.9998)]
    [InlineData(2, 0.5, 3e-7)]
    [InlineData(0.0001, 0.0001, 0.4998779)]
    [InlineData(5, 0.5, 1.230468749943611e-56)]
    [InlineData(5, 0.5, 0.9999)]
    [InlineData(0.5, 5, 7.8e-7)]
    [InlineData(1000, 1000, 0.69)]
    [InlineData(1, 100, 0.894)]
    [InlineData(100, 1, 2.38e-30)]
    [InlineData(0.1, 10, 0.9999999999985647)]
    public void InverseBetaInvertsBeta(double a, double b, double p)
    {
        var (x, y) = Special.InverseBeta(a, b, p, 1 - p);
        Assert.Equal(1, x + y, 15);
        AssertClose(p, Special.BetaRegularized(a, b, x, y), 1e-12);
        var (ux, uy) = Special.InverseBeta(a, b, 1 - p, p);
        AssertClose(p, Special.BetaRegularized(a, b, ux, uy, upper: true), 1e-12);
    }

    [Fact]
    public void StudentTailKeepsPrecisionForManyDegreesOfFreedom()
    {
        // P(T < -2.727679958778845) with 1e7 degrees of freedom: I_x(a, b) with a huge, b = 1/2
        // and x near 1, where the continued fraction alone loses about 1e-10.
        AssertClose(0.003189078390824809, DistributionFunctions.StudentTail(-2.727679958778845, 1e7), 1e-12);
    }

    [Fact]
    public void ElementaryHelpersKeepPrecisionNearZero()
    {
        AssertClose(1e-10 - 5e-21, Special.Log1p(1e-10), 1e-15);
        AssertClose(9.999999999500000e-11, Special.Expm1(-1e-10) * -1, 1e-15);
        AssertClose(-4.999999999666667e-21, Special.Log1pmx(1e-10), 1e-15);
        Assert.Equal(0, Special.SinPi(3));
        Assert.Equal(1, Special.SinPi(0.5), 15);
    }

    private static void AssertClose(double expected, double actual, double tolerance)
    {
        if (expected == 0)
        {
            Assert.Equal(0, actual);
            return;
        }

        var error = Math.Abs(actual - expected) / Math.Abs(expected);
        Assert.True(error <= tolerance, $"expected {expected:R}, got {actual:R} (relative error {error:E2})");
    }
}
