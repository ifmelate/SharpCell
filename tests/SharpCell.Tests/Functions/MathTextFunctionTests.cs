using SharpCell;

namespace SharpCell.Tests.Functions;

public class MathTextFunctionTests
{
    private readonly Workbook _wb = new();

    public MathTextFunctionTests() => _wb.AddSheet("S");

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    [Theory]
    [InlineData("=ROMAN(499)", "CDXCIX")]
    [InlineData("=ROMAN(499,1)", "LDVLIV")]
    [InlineData("=ROMAN(499,2)", "XDIX")]
    [InlineData("=ROMAN(499,3)", "VDIV")]
    [InlineData("=ROMAN(499,4)", "ID")]
    [InlineData("=ROMAN(99,2)", "IC")]
    [InlineData("=ROMAN(3999)", "MMMCMXCIX")]
    [InlineData("=ROMAN(2596,1)", "MMDVCI")]
    [InlineData("=ROMAN(99,TRUE)", "XCIX")]
    [InlineData("=ROMAN(99,FALSE)", "IC")]
    [InlineData("=ROMAN(5.99)", "V")]
    [InlineData("=ROMAN(0)", "")]
    [InlineData("=BASE(16,4,8)", "00000100")]
    [InlineData("=BASE(12345678,16)", "BC614E")]
    [InlineData("=BASE(123456789,36)", "21I3V9")]
    [InlineData("=BASE(12.567,2)", "1100")]
    [InlineData("=BASE(0,2)", "0")]
    public void Writes_numbers_like_Excel(string formula, string expected)
    {
        Assert.Equal(CellValue.Text(expected), Eval(formula));
    }

    [Theory]
    [InlineData("=ARABIC(\"MMMCMXCIX\")", 3999)]
    [InlineData("=ARABIC(\"ixix\")", 18)]
    [InlineData("=ARABIC(\"IM\")", 999)]
    [InlineData("=ARABIC(\"XIXI\")", 20)]
    [InlineData("=ARABIC(\" -MMXI \")", -2011)]
    [InlineData("=ARABIC(\"\")", 0)]
    [InlineData("=DECIMAL(\"FF\",16)", 255)]
    [InlineData("=DECIMAL(\"zz\",36)", 1295)]
    [InlineData("=DECIMAL(111,2)", 7)]
    public void Reads_numbers_like_Excel(string formula, double expected)
    {
        Assert.Equal(CellValue.Number(expected), Eval(formula));
    }

    [Theory]
    [InlineData("=ROMAN(4000)", ErrorKind.Value)]
    [InlineData("=ROMAN(-1)", ErrorKind.Value)]
    [InlineData("=ROMAN(10,5)", ErrorKind.Value)]
    [InlineData("=ARABIC(\"ABC\")", ErrorKind.Value)]
    [InlineData("=ARABIC(10)", ErrorKind.Value)]
    [InlineData("=ARABIC(TRUE)", ErrorKind.Value)]
    [InlineData("=BASE(-1,2)", ErrorKind.Num)]
    [InlineData("=BASE(10,1)", ErrorKind.Num)]
    [InlineData("=BASE(10,37)", ErrorKind.Num)]
    [InlineData("=BASE(10,2,256)", ErrorKind.Num)]
    [InlineData("=BASE(2^53,2)", ErrorKind.Num)]
    [InlineData("=DECIMAL(\"12\",2)", ErrorKind.Num)]
    [InlineData("=DECIMAL(\"1\",37)", ErrorKind.Num)]
    [InlineData("=DECIMAL(1/0,2)", ErrorKind.Div0)]
    public void Fails_like_Excel(string formula, ErrorKind expected)
    {
        Assert.Equal(CellValue.Error(expected), Eval(formula));
    }

    [Fact]
    public void ARABIC_reads_every_ROMAN_form_back()
    {
        for (var form = 0; form <= 4; form++)
        {
            foreach (var n in new[] { 1, 4, 9, 14, 40, 45, 49, 90, 95, 99, 400, 450, 490, 495, 499, 900, 999, 1999, 3999 })
                Assert.Equal(CellValue.Number(n), Eval($"=ARABIC(ROMAN({n},{form}))"));
        }
    }
}
