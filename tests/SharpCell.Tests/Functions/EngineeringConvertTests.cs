using System;
using System.Globalization;
using SharpCell;

namespace SharpCell.Tests.Functions;

public class EngineeringConvertTests
{
    private readonly Workbook _wb = new();

    public EngineeringConvertTests()
    {
        var s = _wb.AddSheet("S");
        s["A1"].Value = true;
    }

    // Values from Excel.
    [Theory]
    [InlineData(1, "lbm", "kg", 0.45359237)]
    [InlineData(1, "sg", "g", 14593.902937206363)]
    [InlineData(0.05, "Mu", "g", 8.30269391e-20)]
    [InlineData(4, "parsec", "m", 1.2342710325126213e+17)]
    [InlineData(1, "Mly", "m", 9.4607304725808e+21)]
    [InlineData(1, "survey_mi", "m", 1609.3472186944373)]
    [InlineData(4, "day", "yr", 0.010951403148528405)]
    [InlineData(1, "ms", "s", 0.001)]
    [InlineData(1, "Torr", "Pa", 133.32236842105263)]
    [InlineData(1, "psi", "Pa", 6894.757293168362)]
    [InlineData(0.05, "MmmHg", "Pa", 6666100)]
    [InlineData(1, "lbf", "N", 4.4482216152605005)]
    [InlineData(1, "HPh", "J", 2684519.537696173)]
    [InlineData(1, "flb", "J", 1.3558179483314003)]
    [InlineData(1, "HP", "W", 745.6998715822702)]
    [InlineData(1, "tbs", "m3", 1.4786764781249999e-05)]
    [InlineData(1, "barrel", "m^3", 0.158987294928)]
    [InlineData(1, "mi3", "m3", 4168181825.4405794)]
    [InlineData(1, "Picapt3", "m3", 4.3903956618655696e-11)]
    [InlineData(0.0002, "Mang3", "m3", 2.0000000000000002e-16)]
    [InlineData(0.0002, "Muk_pt", "m3", 0.11365225000000002)]
    [InlineData(1, "MTON", "m3", 1.13267386368)]
    [InlineData(1, "us_acre", "m2", 4046.8726098742522)]
    [InlineData(0.05, "Mm2", "m2", 50000000000)]
    [InlineData(0.05, "Mar", "m2", 5000000)]
    [InlineData(1, "ly2", "m2", 8.95054210748189e+31)]
    [InlineData(1, "admkn", "m/s", 0.5147733333333333)]
    [InlineData(0.004, "Mmph", "m/s", 1788.16)]
    [InlineData(1, "mbyte", "bit", 0.008)]
    [InlineData(1, "kibyte", "bit", 8192)]
    [InlineData(1, "C", "K", 274.15)]
    [InlineData(1, "F", "K", 255.92777777777775)]
    [InlineData(1, "Rank", "K", 0.5555555555555556)]
    [InlineData(1, "Reau", "K", 274.4)]
    [InlineData(100, "C", "F", 212)]
    [InlineData(12, "MK", "K", 12000000)]
    [InlineData(1, "dam", "m", 10)]
    [InlineData(1, "hPa", "Pa", 100)]
    public void Converts_like_Excel(double number, string from, string to, double expected)
    {
        var actual = _wb.Evaluate($"=CONVERT({number.ToString("R", CultureInfo.InvariantCulture)},\"{from}\",\"{to}\")");
        Assert.Equal(CellValueKind.Number, actual.Kind);
        Assert.True(Math.Abs(actual.AsNumber() - expected) <= 1e-12 * Math.Abs(expected), $"{from}->{to}: expected {expected}, got {actual}");
    }

    [Theory]
    [InlineData("=CONVERT(1,\"Msg\",\"g\")")]
    [InlineData("=CONVERT(1,\"Mmi\",\"m\")")]
    [InlineData("=CONVERT(1,\"MC\",\"K\")")]
    [InlineData("=CONVERT(1,\"Mha\",\"m2\")")]
    [InlineData("=CONVERT(1,\"m\",\"g\")")]
    [InlineData("=CONVERT(1,\"M\",\"m\")")]
    [InlineData("=CONVERT(1,\"KM\",\"m\")")]
    [InlineData("=CONVERT(1,\"kim\",\"m\")")]
    [InlineData("=CONVERT(1,\"\",\"m\")")]
    [InlineData("=CONVERT(1,1,\"m\")")]
    public void Unknown_or_mismatched_units_are_NA(string formula) =>
        Assert.Equal(CellValue.Error(ErrorKind.NA), _wb.Evaluate(formula));

    [Fact]
    public void Number_must_be_a_number()
    {
        Assert.Equal(CellValue.Error(ErrorKind.Value), _wb.Evaluate("=CONVERT(\"x\",\"m\",\"ft\")"));
        Assert.Equal(CellValue.Error(ErrorKind.Value), _wb.Evaluate("=CONVERT(A1,\"m\",\"ft\")"));
        Assert.Equal(CellValue.Error(ErrorKind.Div0), _wb.Evaluate("=CONVERT(1/0,\"m\",\"ft\")"));
        Assert.Equal(CellValue.Number(2000), _wb.Evaluate("=CONVERT(\"2\",\"km\",\"m\")"));
    }
}
