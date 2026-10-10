using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class IterationReaderTests
{
    private static Workbook Load(string calcPr) =>
        XlsxReader.Load(new TestXlsx { WorkbookTail = calcPr }.Sheet("S", "").Build());

    [Fact]
    public void Iteration_settings_come_from_calcPr()
    {
        Assert.Equal(new IterationSettings(true, 7, 0.5), Load("<calcPr iterate=\"1\" iterateCount=\"7\" iterateDelta=\"0.5\"/>").Iteration);
        Assert.Equal(new IterationSettings(), Load("").Iteration);
        Assert.Equal(new IterationSettings(false, 7, 0.001), Load("<calcPr iterateCount=\"7\"/>").Iteration);
        Assert.Equal(new IterationSettings(true, 100, 0.001), Load("<calcPr iterate=\"true\"/>").Iteration);
    }

    [Theory]
    [InlineData("<calcPr iterateCount=\"x\"/>")]
    [InlineData("<calcPr iterateCount=\"0\"/>")]
    [InlineData("<calcPr iterateDelta=\"-1\"/>")]
    [InlineData("<calcPr iterateDelta=\"abc\"/>")]
    public void Bad_calcPr_numbers_are_invalid_data(string calcPr)
    {
        Assert.Throws<InvalidDataException>(() => Load(calcPr));
    }

    [Fact]
    public void Changed_iteration_settings_cannot_be_saved()
    {
        var workbook = XlsxWriterTests.Load(new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"));
        workbook.Iteration = new IterationSettings(Enabled: true);
        workbook.Recalculate();
        Assert.Throws<NotSupportedException>(() => XlsxWriterTests.Save(workbook));
    }
}
