using SharpCell;

namespace SharpCell.Tests.Functions;

public class EngineeringFunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public EngineeringFunctionTests()
    {
        _s = _wb.AddSheet("S");
        _s["A1"].Value = true;
        _s["A2"].Value = "12";
        _s["A3"].Value = 1100100;
    }

    private CellValue Eval(string formula) => _wb.Evaluate(formula);

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    public static TheoryData<string, CellValue> DeltaAndGestep => new()
    {
        { "=DELTA(3,3)", 1 },
        { "=DELTA(3,3.000000000000001)", 0 },
        { "=DELTA(0)", 1 },
        { "=DELTA(\"2\",2)", 1 },
        { "=DELTA(A1,1)", Err(ErrorKind.Value) },
        { "=DELTA(TRUE)", Err(ErrorKind.Value) },
        { "=DELTA(\"x\")", Err(ErrorKind.Value) },
        { "=DELTA(1/0,1)", Err(ErrorKind.Div0) },
        { "=GESTEP(-3,-3)", 1 },
        { "=GESTEP(-2,-7)", 1 },
        { "=GESTEP(-7,-2)", 0 },
        { "=GESTEP(-1)", 0 },
        { "=GESTEP(A2)", 1 },
        { "=GESTEP(A1)", Err(ErrorKind.Value) },
    };

    [Theory]
    [MemberData(nameof(DeltaAndGestep))]
    public void Delta_and_gestep(string formula, CellValue expected) => Assert.Equal(expected, Eval(formula));

    public static TheoryData<string, CellValue> Bits => new()
    {
        { "=BITAND(13,25)", 9 },
        { "=BITOR(13,25)", 29 },
        { "=BITXOR(13,25)", 20 },
        { "=BITAND(2^48-1,1)", 1 },
        { "=BITOR(2^48-1,1)", 281474976710655 },
        { "=BITAND(2^48,1)", Err(ErrorKind.Num) },
        { "=BITAND(-1,1)", Err(ErrorKind.Num) },
        { "=BITAND(12.5,1)", Err(ErrorKind.Num) },
        { "=BITOR(A1,4)", 5 },
        { "=BITAND(\"123\",456)", 72 },
        { "=BITAND(\"x\",456)", Err(ErrorKind.Value) },
        { "=BITLSHIFT(12,40)", 13194139533312 },
        { "=BITLSHIFT(12,45)", Err(ErrorKind.Num) },
        { "=BITLSHIFT(12,-2)", 3 },
        { "=BITLSHIFT(12,-53)", 0 },
        { "=BITLSHIFT(12,-54)", Err(ErrorKind.Num) },
        { "=BITLSHIFT(12,34.67)", 206158430208 },
        { "=BITRSHIFT(12,2)", 3 },
        { "=BITRSHIFT(145,-3)", 1160 },
        { "=BITRSHIFT(12,50)", 0 },
        { "=BITRSHIFT(12,54)", Err(ErrorKind.Num) },
        { "=BITRSHIFT(12.6,1)", Err(ErrorKind.Num) },
    };

    [Theory]
    [MemberData(nameof(Bits))]
    public void Bit_functions(string formula, CellValue expected) => Assert.Equal(expected, Eval(formula));

    public static TheoryData<string, CellValue> FromDecimal => new()
    {
        { "=DEC2BIN(9)", "1001" },
        { "=DEC2BIN(9,8)", "00001001" },
        { "=DEC2BIN(511,8)", Err(ErrorKind.Num) },
        { "=DEC2BIN(511.9)", "111111111" },
        { "=DEC2BIN(512)", Err(ErrorKind.Num) },
        { "=DEC2BIN(-512.8)", "1000000000" },
        { "=DEC2BIN(-513)", Err(ErrorKind.Num) },
        { "=DEC2BIN(-1,3)", "1111111111" },
        { "=DEC2BIN(-1,11)", Err(ErrorKind.Num) },
        { "=DEC2BIN(1,0)", Err(ErrorKind.Num) },
        { "=DEC2BIN(1,2.7)", "01" },
        { "=DEC2BIN(1,TRUE)", Err(ErrorKind.Value) },
        { "=DEC2BIN(A2)", "1100" },
        { "=DEC2BIN(A1)", Err(ErrorKind.Value) },
        { "=DEC2BIN(0,2)", "00" },
        { "=DEC2OCT(-1213)", "7777775503" },
        { "=DEC2OCT(536870911)", "3777777777" },
        { "=DEC2OCT(536870912)", Err(ErrorKind.Num) },
        { "=DEC2HEX(121233)", "1D991" },
        { "=DEC2HEX(549755813887)", "7FFFFFFFFF" },
        { "=DEC2HEX(549755813888)", Err(ErrorKind.Num) },
        { "=DEC2HEX(-549755813888)", "8000000000" },
        { "=DEC2HEX(-1213)", "FFFFFFFB43" },
    };

    [Theory]
    [MemberData(nameof(FromDecimal))]
    public void Decimal_to_other_bases(string formula, CellValue expected) => Assert.Equal(expected, Eval(formula));

    public static TheoryData<string, CellValue> ToOtherBases => new()
    {
        { "=BIN2DEC(1100100)", 100 },
        { "=BIN2DEC(\"1110111011\")", -69 },
        { "=BIN2DEC(1000000000)", -512 },
        { "=BIN2DEC(10000000000)", Err(ErrorKind.Num) },
        { "=BIN2DEC(3)", Err(ErrorKind.Num) },
        { "=BIN2DEC(-1)", Err(ErrorKind.Num) },
        { "=BIN2DEC(A1)", Err(ErrorKind.Value) },
        { "=BIN2DEC(Z99)", 0 },
        { "=BIN2HEX(A3,3)", "064" },
        { "=BIN2HEX(1110111011,4)", "FFFFFFFFBB" },
        { "=BIN2HEX(1110111011,11)", Err(ErrorKind.Num) },
        { "=BIN2OCT(1000000000)", "7777777000" },
        { "=BIN2OCT(\"0000001000\")", "10" },
        { "=HEX2DEC(\"01fffffff\")", 536870911 },
        { "=HEX2DEC(8000000000)", -549755813888 },
        { "=HEX2DEC(9999999999)", -439804651111 },
        { "=HEX2DEC(\"QWERTY\")", Err(ErrorKind.Num) },
        { "=HEX2BIN(\"FE\",3)", Err(ErrorKind.Num) },
        { "=HEX2BIN(\"FE\",9)", "011111110" },
        { "=HEX2BIN(\"FFFFFFFE00\")", "1000000000" },
        { "=HEX2BIN(\"FFEF\")", Err(ErrorKind.Num) },
        { "=HEX2OCT(\"FFFFFFFF\")", Err(ErrorKind.Num) },
        { "=HEX2OCT(\"FFE0000000\")", "4000000000" },
        { "=OCT2DEC(4000000000)", -536870912 },
        { "=OCT2DEC(3777777778)", Err(ErrorKind.Num) },
        { "=OCT2HEX(7000000000,2)", "FFF8000000" },
        { "=OCT2HEX(1234566,10)", "0000053976" },
        { "=OCT2BIN(776)", "111111110" },
        { "=OCT2BIN(7777777777)", "1111111111" },
        { "=OCT2BIN(1000)", Err(ErrorKind.Num) },
    };

    [Theory]
    [MemberData(nameof(ToOtherBases))]
    public void Conversions_between_bases(string formula, CellValue expected) => Assert.Equal(expected, Eval(formula));

    [Fact]
    public void Base_conversions_are_applied_element_wise()
    {
        Assert.Equal(CellValue.Array(new CellValue[,] { { "1", "10", "11" } }), Eval("=DEC2BIN({1,2,3})"));
    }
}
