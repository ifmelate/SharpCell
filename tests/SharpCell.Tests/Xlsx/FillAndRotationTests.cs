using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>Pattern and gradient fills, turned, stacked and shrunk text: read, saved and read back as Excel has them.</summary>
public class FillAndRotationTests
{
    private const string Styles = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
          <fills count="6">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="gray125"/></fill>
            <fill><patternFill patternType="darkTrellis"><fgColor rgb="FFFF0000"/><bgColor rgb="FF00FF00"/></patternFill></fill>
            <fill><patternFill patternType="lightGray"><fgColor indexed="64"/><bgColor indexed="65"/></patternFill></fill>
            <fill><gradientFill degree="90"><stop position="0"><color theme="0"/></stop><stop position="1"><color theme="4"/></stop></gradientFill></fill>
            <fill><gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5"><stop position="0"><color rgb="FFFFFFFF"/></stop><stop position="0.5"><color rgb="FF808080"/></stop><stop position="1"><color rgb="FF000000"/></stop></gradientFill></fill>
          </fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellXfs count="10">
            <xf fontId="0" fillId="0" borderId="0"/>
            <xf fontId="0" fillId="2" borderId="0" applyFill="1"/>
            <xf fontId="0" fillId="3" borderId="0" applyFill="1"/>
            <xf fontId="0" fillId="4" borderId="0" applyFill="1"/>
            <xf fontId="0" fillId="5" borderId="0" applyFill="1"/>
            <xf fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment textRotation="45"/></xf>
            <xf fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment textRotation="135"/></xf>
            <xf fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment textRotation="180"/></xf>
            <xf fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment textRotation="255"/></xf>
            <xf fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment shrinkToFit="1"/></xf>
          </cellXfs>
        </styleSheet>
        """;

    private static (byte[] Original, Workbook Workbook) Load()
    {
        var cells = string.Concat(Enumerable.Range(0, 10).Select(i => $"<c r=\"{(char)('A' + i)}1\" s=\"{i}\"><v>1</v></c>"));
        var original = new TestXlsx { Styles = Styles }.RawSheet("S", $"<sheetData><row r=\"1\">{cells}</row></sheetData>").Build().ToArray();
        return (original, XlsxReader.Load(new MemoryStream(original)));
    }

    [Fact]
    public void Patterns_are_drawn_in_their_colour_over_the_background()
    {
        var (_, workbook) = Load();
        var sheet = workbook["S"];

        var trellis = sheet["B1"].Style;
        Assert.Equal((CellFillPattern.DarkTrellis, CellColor.FromRgb(0xFF0000), CellColor.FromRgb(0x00FF00)), (trellis.FillPattern, trellis.PatternColor, trellis.Fill));
        var automatic = sheet["C1"].Style;
        Assert.Equal((CellFillPattern.LightGray, (CellColor?)null, (CellColor?)null), (automatic.FillPattern, automatic.PatternColor, automatic.Fill));
        Assert.Equal((CellFillPattern.Solid, (CellColor?)null), (sheet["A1"].Style.FillPattern, sheet["A1"].Style.Fill));
    }

    [Fact]
    public void Gradients_are_read_with_their_stops()
    {
        var (_, workbook) = Load();

        var linear = workbook["S"]["D1"].Style.Gradient!;
        Assert.Equal((CellGradientType.Linear, 90.0), (linear.Type, linear.Angle));
        Assert.Equal([workbook.ThemeColor(ThemeColor.Background1), workbook.ThemeColor(ThemeColor.Accent1)], linear.Stops.Select(stop => stop.Color));
        var path = workbook["S"]["E1"].Style.Gradient!;
        Assert.Equal((CellGradientType.Path, 0.5, 0.5, 3), (path.Type, path.Left, path.Bottom, path.Stops.Count));
        Assert.Null(workbook["S"]["D1"].Style.Fill);
    }

    [Fact]
    public void Turned_stacked_and_shrunk_text_is_read()
    {
        var (_, workbook) = Load();
        var sheet = workbook["S"];

        Assert.Equal([45, -45, -90, 0, 0], new[] { "F1", "G1", "H1", "I1", "J1" }.Select(a => sheet[a].Style.TextRotation));
        Assert.True(sheet["I1"].Style.VerticalText);
        Assert.True(sheet["J1"].Style.ShrinkToFit);
        Assert.False(sheet["F1"].Style.ShrinkToFit);
    }

    [Fact]
    public void Unchanged_fills_and_alignment_save_as_the_file_has_them()
    {
        var (original, workbook) = Load();
        workbook["S"]["A1"].Value = 2;

        var saved = XlsxWriterTests.Save(workbook);

        Assert.Equal(XlsxWriterTests.Part(original, "xl/styles.xml"), XlsxWriterTests.Part(saved, "xl/styles.xml"));
    }

    [Fact]
    public void New_fills_and_alignment_are_saved_and_read_back()
    {
        var (original, workbook) = Load();
        var sheet = workbook["S"];
        var accent = workbook.ThemeColor(ThemeColor.Accent2, -0.25);
        sheet["A2"].Style = CellStyle.Default with { FillPattern = CellFillPattern.LightUp, PatternColor = accent, Fill = CellColor.FromRgb(0xFFFFCC) };
        sheet["B2"].Style = CellStyle.Default with { Gradient = CellGradient.Linear(CellColor.FromRgb(0xFFFFFF), accent, 45) };
        sheet["C2"].Style = CellStyle.Default with
        {
            Gradient = new CellGradient([new GradientStop(1, CellColor.FromRgb(0)), new GradientStop(0, CellColor.FromRgb(0xFFFFFF))])
            { Type = CellGradientType.Path, Left = 0.5, Right = 0.5, Top = 0.5, Bottom = 0.5 },
        };
        sheet["D2"].Style = CellStyle.Default with { TextRotation = -30, ShrinkToFit = true };
        sheet["E2"].Style = CellStyle.Default with { VerticalText = true };
        sheet["F2"].Style = CellStyle.Default with { FillPattern = CellFillPattern.Gray125 };
        sheet["B1"].Style = sheet["B1"].Style with { PatternColor = CellColor.FromRgb(0x0000FF) };

        var saved = XlsxWriterTests.Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var reloaded = XlsxWriterTests.Reload(saved)["S"];
        var styles = XlsxWriterTests.Part(saved, "xl/styles.xml");

        foreach (var address in new[] { "A2", "B2", "C2", "D2", "E2", "F2", "B1" })
            Assert.Equal(sheet[address].Style, reloaded[address].Style);
        Assert.Contains("<alignment textRotation=\"120\" shrinkToFit=\"1\" />", styles);
        Assert.Contains("<alignment textRotation=\"255\" />", styles);
        Assert.Contains("<gradientFill degree=\"45\">", styles);
        Assert.Contains("<fgColor theme=\"5\" tint=\"-0.25\" />", styles);
        Assert.Equal(2, styles.Split("patternType=\"gray125\"").Length - 1);   // the reserved fill 1 is not reused
    }

    [Fact]
    public void Wrong_values_are_refused()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => CellStyle.Default with { TextRotation = 91 });
        Assert.Throws<ArgumentException>(() => new CellGradient([new GradientStop(0, default)]));
        Assert.Throws<ArgumentOutOfRangeException>(() => new GradientStop(1.5, default));
        Assert.Throws<ArgumentOutOfRangeException>(() => CellGradient.Linear(default, default) with { Left = -0.1 });
    }

    [Fact]
    public void Gradients_compare_by_their_stops()
    {
        var a = CellGradient.Linear(CellColor.FromRgb(1), CellColor.FromRgb(2), 90);
        var b = CellGradient.Linear(CellColor.FromRgb(1), CellColor.FromRgb(2), 90);

        Assert.Equal(a, b);
        Assert.Equal(a.GetHashCode(), b.GetHashCode());
        Assert.NotEqual(a, CellGradient.Linear(CellColor.FromRgb(1), CellColor.FromRgb(3), 90));
    }
}
