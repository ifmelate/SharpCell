using System;
using SharpCell;

namespace SharpCell.Tests.Functions;

public class EngineeringComplexTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public EngineeringComplexTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = "3+4i";
        _s["A2"].Value = "1-2i";
        _s["A3"].Value = 2;
        _s["B1"].Value = true;
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    public static TheoryData<string, CellValue> Texts => new()
    {
        { "=COMPLEX(12,6)", "12+6i" },
        { "=COMPLEX(1,-2,\"j\")", "1-2j" },
        { "=COMPLEX(1,1)", "1+i" },
        { "=COMPLEX(0,-1)", "-i" },
        { "=COMPLEX(0,5)", "5i" },
        { "=COMPLEX(5,0)", "5" },
        { "=COMPLEX(0,0)", "0" },
        { "=COMPLEX(0,1,\"j\")", "j" },
        { "=COMPLEX(1,2,\"\")", "1+2i" },
        { "=COMPLEX(1,2,Z9)", "1+2i" },
        { "=COMPLEX(1,2,0)", Err(ErrorKind.Value) },
        { "=COMPLEX(1,2,\"I\")", Err(ErrorKind.Value) },
        { "=COMPLEX(B1,2)", Err(ErrorKind.Value) },
        { "=COMPLEX(\"1\",2)", "1+2i" },
        { "=COMPLEX(1.2E-45,-3)", "1.2E-45-3i" },
        { "=COMPLEX(1E-16,0)", "0.0000000000000001" },
        { "=COMPLEX(1E-18,0)", "0.000000000000000001" },
        { "=COMPLEX(1E-20,0)", "1E-20" },
        { "=COMPLEX(1/3,0)", "0.333333333333333" },
        { "=COMPLEX(1E+20,1)", "1E+20+i" },
        { "=IMCONJUGATE(\"1-2j\")", "1+2j" },
        { "=IMCONJUGATE(\"-5i\")", "5i" },
        { "=IMCONJUGATE(A3)", "2" },
        { "=IMSUM(A1,\"-4+5i\")", "-1+9i" },
        { "=IMSUM(A1:A3)", "6+2i" },
        { "=IMSUM({\"i\",\"i\"})", "2i" },
        { "=IMSUM(\"i\",\"j\")", Err(ErrorKind.Value) },
        { "=IMSUM(Z1:Z5)", "0" },
        { "=IMSUB(A1,\"-4+5i\")", "7-i" },
        { "=IMSUB(7,Z9)", "7" },
        { "=IMPRODUCT(A1,\"-4+5i\")", "-32-i" },
        { "=IMPRODUCT(\"i\",\"i\")", "-1" },
        { "=IMPRODUCT(3,\"3j\")", "9j" },
        { "=IMPRODUCT(7,Z9)", "0" },
        { "=IMPRODUCT(A1:A4)", "0" },
        { "=IMPRODUCT(A1:A3)", "22-4i" },
        { "=IMPRODUCT(A1,B1)", Err(ErrorKind.Value) },
        { "=IMDIV(\"-238+240i\",\"10+24i\")", "5+12i" },
        { "=IMDIV(1,0)", Err(ErrorKind.Num) },
        { "=IMSQRT(\"-4\")", "1.22464679914735E-16+2i" },
        { "=IMPOWER(\"2+3i\",3)", "-46+9.00000000000001i" },
        { "=IMPOWER(0,2)", "0" },
        { "=IMPOWER(0,0)", Err(ErrorKind.Num) },
        { "=IMEXP(0)", "1" },
        { "=IMEXP(1000)", Err(ErrorKind.Num) },
        { "=IMLN(0)", Err(ErrorKind.Num) },
        { "=IMCOT(0)", Err(ErrorKind.Num) },
        { "=IMCSC(0)", Err(ErrorKind.Num) },
        { "=IMCSCH(0)", Err(ErrorKind.Num) },
        { "=IMSEC(0)", "1" },
        { "=IMSECH(0)", "1" },
        { "=IMSIN(\"x\")", Err(ErrorKind.Num) },
        { "=IMSIN(\"3 + 4i\")", Err(ErrorKind.Num) },
        { "=IMSIN(\"3+4ij\")", Err(ErrorKind.Num) },
        { "=IMSIN(1/0)", Err(ErrorKind.Div0) },
    };

    [Theory]
    [MemberData(nameof(Texts))]
    public void Complex_results(string formula, CellValue expected) => Assert.Equal(expected, Eval(formula));

    public static TheoryData<string, CellValue> Numbers => new()
    {
        { "=IMREAL(\"-23.1234-45.899i\")", -23.1234 },
        { "=IMAGINARY(\"-23.1234-45.899i\")", -45.899 },
        { "=IMAGINARY(\"i\")", 1 },
        { "=IMAGINARY(\"-j\")", -1 },
        { "=IMAGINARY(\"+i\")", 1 },
        { "=IMREAL(\"1.2E-45-3i\")", 1.2e-45 },
        { "=IMAGINARY(\"1E-5i\")", 1e-5 },
        { "=IMAGINARY(\"4.5E+2i\")", 450 },
        { "=IMREAL(5)", 5 },
        { "=IMREAL(B1)", Err(ErrorKind.Value) },
        { "=IMREAL(\"3+4\")", Err(ErrorKind.Num) },
        { "=IMREAL(\"3+-4i\")", Err(ErrorKind.Num) },
        { "=IMABS(\"3+4i\")", 5 },
        { "=IMARGUMENT(\"-1\")", Math.PI },
        { "=IMARGUMENT(0)", Err(ErrorKind.Div0) },
    };

    [Theory]
    [MemberData(nameof(Numbers))]
    public void Complex_parts(string formula, CellValue expected) => Assert.Equal(expected, Eval(formula));

    // Values from Excel; results are text with 15 significant digits.
    [Theory]
    [InlineData("=IMCOS(\"12+6i\")", 170.218538080177, 108.233817449252)]
    [InlineData("=IMSIN(\"12+6i\")", -108.23514748054, 170.21644637534)]
    [InlineData("=IMTAN(\"12+6i\")", -1.11280735120837e-05, 0.999994787459874)]
    [InlineData("=IMSINH(\"12+6i\")", 78136.157362303, -22738.10556516)]
    [InlineData("=IMCOSH(\"12+6i\")", 78136.1573682025, -22738.1055634432)]
    [InlineData("=IMCOT(\"12+6i\")", -1.11281895226721e-05, -1.00000521244346)]
    [InlineData("=IMCSC(\"12+6i\")", -0.00266009278646834, -0.00418340577604633)]
    [InlineData("=IMCSCH(\"12+6i\")", 1.17989802733781e-05, 3.43357631183878e-06)]
    [InlineData("=IMSEC(\"12+6i\")", 0.00418341357158393, -0.00266003236738094)]
    [InlineData("=IMSECH(\"12+6i\")", 1.17989802727654e-05, 3.43357631114201e-06)]
    [InlineData("=IMLN(\"12+6i\")", 2.59647842544511, 0.463647609000806)]
    [InlineData("=IMLOG10(\"12+6i\")", 1.12763625255165, 0.201359598136687)]
    [InlineData("=IMLOG2(\"12+6i\")", 3.74592654816484, 0.668902106225488)]
    [InlineData("=IMEXP(\"12+6i\")", 156272.314730506, -45476.2111286032)]
    [InlineData("=IMSQRT(\"12+6i\")", 3.56485678990045, 0.841548532468193)]
    [InlineData("=IMSINH(\"5i\")", 0, -0.958924274663138)]
    [InlineData("=IMCOT(\"1E-16\")", 1e16, 0)]
    [InlineData("=IMCSC(\"1E-17\")", 1e17, 0)]
    public void Complex_functions_match_Excel(string formula, double real, double imaginary)
    {
        var inner = formula[1..];
        AssertClose(real, Eval("=IMREAL(" + inner + ")"));
        AssertClose(imaginary, Eval("=IMAGINARY(" + inner + ")"));
    }

    [Fact]
    public void Parts_use_the_workbook_decimal_separator()
    {
        _wb.Culture = new System.Globalization.CultureInfo("de-DE");
        Assert.Equal(CellValue.Text("1,5+2i"), Eval("=COMPLEX(1.5,2)"));
        Assert.Equal(CellValue.Number(2.5), Eval("=IMREAL(\"2,5-i\")"));
        Assert.Equal(CellValue.Number(2.5), Eval("=IMREAL(\"2.5-i\")"));
    }

    [Theory]
    [InlineData("=IMREAL(IMTAN(\"7-12.6i\"))", 2.25273756034272E-11)]
    [InlineData("=IMREAL(IMCOT(\"7-12.6i\"))", 2.25273756035673E-11)]
    [InlineData("=IMREAL(IMTAN(\"-23.1234-45.899i\"))", -2.0870098138858E-40)]
    public void Tangents_keep_small_real_parts(string formula, double expected) => AssertClose(expected, Eval(formula), 1e-13);

    private static void AssertClose(double expected, CellValue actual, double relative)
    {
        Assert.Equal(CellValueKind.Number, actual.Kind);
        Assert.True(Math.Abs(expected - actual.AsNumber()) <= relative * Math.Abs(expected), $"expected {expected}, got {actual}");
    }

    private static void AssertClose(double expected, CellValue actual)
    {
        Assert.Equal(CellValueKind.Number, actual.Kind);
        Assert.True(Math.Abs(expected - actual.AsNumber()) <= 1e-13 * Math.Max(1, Math.Abs(expected)),
            $"expected {expected}, got {actual}");
    }
}
