using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>Named cell styles (Good, Heading 1, custom ones): read, given to cells, added, changed and removed.</summary>
public class NamedStyleTests
{
    private const string Styles = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/></font></fonts>
          <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill></fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/></cellStyleXfs>
          <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="1" applyFill="1"/></cellXfs>
          <cellStyles count="2"><cellStyle name="Good" xfId="1" builtinId="26"/><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>
        """;

    private static (byte[] Original, Workbook Workbook) Load()
    {
        var original = new TestXlsx { Styles = Styles }.RawSheet("S", """
            <sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>2</v></c><c r="C1"><v>3</v></c></row></sheetData>
            """).Build().ToArray();
        return (original, XlsxReader.Load(new MemoryStream(original)));
    }

    private static readonly CellStyle GoodLook = CellStyle.Default with
    {
        Font = CellFont.Default with { Color = CellColor.FromRgb(0x006100) },
        Fill = CellColor.FromRgb(0xC6EFCE),
    };

    [Fact]
    public void Named_styles_and_the_styles_of_cells_based_on_them_are_read()
    {
        var (_, workbook) = Load();
        var sheet = workbook["S"];

        var good = Assert.Single(workbook.NamedStyles).Value;
        Assert.Equal(GoodLook with { Name = "Good" }, good);
        Assert.Same(good, workbook.NamedStyles["good"]);
        Assert.Equal("Good", sheet["A1"].Style.Name);
        Assert.Equal(("Good", (CellColor?)null), (sheet["B1"].Style.Name, sheet["B1"].Style.Fill));   // Good with its fill taken off
        Assert.Null(sheet["C1"].Style.Name);
    }

    [Fact]
    public void Normal_is_no_name_and_names_cannot_be_empty()
    {
        Assert.Null((CellStyle.Default with { Name = "normal" }).Name);
        Assert.Throws<ArgumentException>(() => CellStyle.Default with { Name = " " });
        Assert.Throws<ArgumentException>(() => new Workbook().SetNamedStyle("Normal", CellStyle.Default));
    }

    [Fact]
    public void Unchanged_named_styles_save_as_the_file_has_them()
    {
        var (original, workbook) = Load();
        workbook["S"]["C1"].Value = 4;

        var saved = XlsxWriterTests.Save(workbook);

        Assert.Equal(XlsxWriterTests.Part(original, "xl/styles.xml"), XlsxWriterTests.Part(saved, "xl/styles.xml"));
    }

    [Fact]
    public void A_named_style_given_to_a_cell_is_saved_with_its_name()
    {
        var (original, workbook) = Load();
        var sheet = workbook["S"];
        var heading = workbook.SetNamedStyle("Heading 1", CellStyle.Default with { Font = CellFont.Default with { Bold = true, Size = 15 } });
        var mine = workbook.SetNamedStyle("My Total", CellStyle.Default with { TopBorder = new CellBorder(CellBorderStyle.Double, null) });
        sheet["C1"].Style = heading;
        sheet["D1"].Style = mine;
        sheet["E1"].Style = workbook.NamedStyles["Good"];
        sheet["A1"].Style = sheet["A1"].Style with { Name = null };

        var saved = XlsxWriterTests.Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var styles = XlsxWriterTests.Part(saved, "xl/styles.xml");
        var reloaded = XlsxWriterTests.Reload(saved);

        Assert.Contains("<cellStyle name=\"Heading 1\" xfId=\"2\" builtinId=\"16\" customBuiltin=\"1\" />", styles);
        Assert.Contains("<cellStyle name=\"My Total\" xfId=\"3\" />", styles);
        Assert.Equal(["Good", "Heading 1", "My Total"], reloaded.NamedStyles.Keys.Order());
        Assert.Equal(heading, reloaded["S"]["C1"].Style);
        Assert.Equal(mine, reloaded["S"]["D1"].Style);
        Assert.Equal(workbook.NamedStyles["Good"], reloaded["S"]["E1"].Style);
        Assert.Equal(GoodLook, reloaded["S"]["A1"].Style);
    }

    [Fact]
    public void A_changed_named_style_saves_its_new_look_and_a_removed_one_goes()
    {
        var (original, workbook) = Load();
        workbook.SetNamedStyle("Good", GoodLook with { Font = GoodLook.Font with { Bold = true } });

        var saved = XlsxWriterTests.Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        Assert.True(XlsxWriterTests.Reload(saved).NamedStyles["Good"].Font.Bold);

        var reloaded = XlsxWriterTests.Reload(saved);
        Assert.True(reloaded.RemoveNamedStyle("Good"));
        Assert.Null(reloaded["S"]["A1"].Style.Name);
        Assert.Equal(CellStyle.Default, reloaded["S"]["A1"].Style with { });
        var again = XlsxWriterTests.Save(reloaded);
        Assert.DoesNotContain("name=\"Good\"", XlsxWriterTests.Part(again, "xl/styles.xml"));
        Assert.Empty(XlsxWriterTests.Reload(again).NamedStyles);
    }

    [Fact]
    public void A_workbook_built_in_code_saves_its_named_styles()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = "ok";
        sheet["A1"].Style = workbook.SetNamedStyle("Good", GoodLook);
        workbook.SetNamedStyle("Unused", CellStyle.Default with { WrapText = true });

        var saved = XlsxWriterTests.Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        var reloaded = XlsxWriterTests.Reload(saved);
        Assert.Equal(workbook.NamedStyles["Good"], reloaded["S"]["A1"].Style);
        Assert.True(reloaded.NamedStyles["Unused"].WrapText);
        Assert.Equal(workbook.NamedStyles["Good"], workbook.Clone().NamedStyles["Good"]);
    }
}
