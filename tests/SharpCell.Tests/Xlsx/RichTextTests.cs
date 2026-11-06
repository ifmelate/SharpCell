using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>Text in runs of different fonts: read from shared and inline strings, set in code, saved and read back.</summary>
public class RichTextTests
{
    private static readonly CellFont Bold = CellFont.Default with { Bold = true };

    [Fact]
    public void Rich_text_sets_the_value_and_a_new_value_or_formula_takes_the_runs_away()
    {
        var sheet = new Workbook().AddSheet("S");
        var rich = new RichText(new TextRun("Total "), new TextRun("", Bold), new TextRun("m"), new TextRun("2", CellFont.Default with { Script = CellFontScript.Superscript }));

        sheet["A1"].RichText = rich;

        Assert.Equal("Total m2", sheet["A1"].Value.AsText());
        Assert.Equal(3, sheet["A1"].RichText!.Runs.Count);   // the empty run is dropped
        sheet["A1"].Value = "plain";
        Assert.Null(sheet["A1"].RichText);
        sheet["A2"].RichText = rich;
        sheet["A2"].Formula = "=\"x\"";
        Assert.Null(sheet["A2"].RichText);
        sheet["A3"].RichText = rich;
        sheet["A3"].RichText = null;
        Assert.Equal(("Total m2", (RichText?)null), (sheet["A3"].Value.AsText(), sheet["A3"].RichText));
    }

    [Fact]
    public void Wrong_rich_text_is_refused_and_equal_runs_compare_equal()
    {
        Assert.Throws<ArgumentException>(() => new RichText(new TextRun("")));
        Assert.Throws<ArgumentException>(() => new RichText(new TextRun(new string('x', 32767)), new TextRun("y")));
        Assert.Equal(new RichText(new TextRun("a", Bold), new TextRun("b")), new RichText(new TextRun("a", Bold), new TextRun("b")));
        Assert.NotEqual(new RichText(new TextRun("a", Bold), new TextRun("b")), new RichText(new TextRun("a"), new TextRun("b")));
    }

    [Fact]
    public void Runs_move_with_their_cells_and_are_copied_and_cloned()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        var rich = new RichText(new TextRun("a", Bold), new TextRun("b"));
        sheet["A1"].RichText = rich;

        sheet.InsertRows(1);
        sheet.Range("A2").CopyTo(sheet.Range("C5"));

        Assert.Equal(rich, sheet["A2"].RichText);
        Assert.Equal(rich, sheet["C5"].RichText);
        Assert.Equal(rich, workbook.Clone()["S"]["A2"].RichText);
    }

    private const string Shared = """
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
          <si><t>plain</t></si>
          <si><r><t xml:space="preserve">Hello </t></r><r><rPr><b/><sz val="11"/><color theme="4"/><rFont val="Calibri"/></rPr><t>world</t></r><rPh sb="0" eb="1"><t>ハ</t></rPh></si>
          <si><r><rPr><i/><sz val="9"/><rFont val="Arial"/></rPr><t>only</t></r></si>
        </sst>
        """;

    private static (byte[] Original, Workbook Workbook) Load()
    {
        var original = new TestXlsx { SharedStrings = Shared }.RawSheet("S", """
            <sheetData><row r="1">
              <c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c>
              <c r="D1" t="inlineStr"><is><r><rPr><strike/><vertAlign val="subscript"/><sz val="11"/><rFont val="Calibri"/></rPr><t>H</t></r><r><t>2O</t></r></is></c>
            </row></sheetData>
            """).Build().ToArray();
        return (original, XlsxReader.Load(new MemoryStream(original)));
    }

    [Fact]
    public void Runs_are_read_from_shared_and_inline_strings()
    {
        var (_, workbook) = Load();
        var sheet = workbook["S"];

        Assert.Null(sheet["A1"].RichText);
        var hello = sheet["B1"].RichText!;
        Assert.Equal("Hello world", sheet["B1"].Value.AsText());   // the phonetic guide is not text
        Assert.Equal(new TextRun("Hello "), hello.Runs[0]);   // no rPr: the cell's font
        Assert.Equal(CellFont.Default with { Bold = true, Color = workbook.ThemeColor(ThemeColor.Accent1) }, hello.Runs[1].Font);
        Assert.Equal(CellFont.Default with { Italic = true, Size = 9, Name = "Arial" }, sheet["C1"].RichText!.Runs.Single().Font);
        var water = sheet["D1"].RichText!;
        Assert.Equal(("H2O", CellFontScript.Subscript, true), (water.Text, water.Runs[0].Font!.Script, water.Runs[0].Font!.Strikethrough));
    }

    [Fact]
    public void Unchanged_rich_text_saves_as_the_file_has_it()
    {
        var (original, workbook) = Load();
        workbook["S"]["E1"].Value = 1;

        var part = XlsxWriterTests.Part(XlsxWriterTests.Save(workbook), "xl/worksheets/sheet1.xml");

        Assert.Contains("<c r=\"B1\" t=\"s\"><v>1</v></c>", part);
        Assert.Contains("<rFont val=\"Calibri\" /></rPr><t>H</t>", part);
    }

    [Fact]
    public void New_and_changed_rich_text_is_saved_and_read_back()
    {
        var (original, workbook) = Load();
        var sheet = workbook["S"];
        var accent = workbook.ThemeColor(ThemeColor.Accent2);
        sheet["A1"].RichText = new RichText(new TextRun("  lead "), new TextRun("bold", Bold), new TextRun(" & <tail>  ", CellFont.Default with { Color = accent }));
        sheet["B1"].RichText = null;
        sheet["F1"].RichText = new RichText(new TextRun("x"), new TextRun("2", CellFont.Default with { Script = CellFontScript.Superscript }));

        var saved = XlsxWriterTests.Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var reloaded = XlsxWriterTests.Reload(saved)["S"];
        var part = XlsxWriterTests.Part(saved, "xl/worksheets/sheet1.xml");

        Assert.Equal(sheet["A1"].RichText, reloaded["A1"].RichText);
        Assert.Equal(("Hello world", (RichText?)null), (reloaded["B1"].Value.AsText(), reloaded["B1"].RichText));
        Assert.Equal(sheet["F1"].RichText, reloaded["F1"].RichText);
        Assert.Equal(sheet["C1"].RichText, reloaded["C1"].RichText);
        Assert.Contains("<t xml:space=\"preserve\">  lead </t>", part);
        Assert.DoesNotContain("<r><t xml:space=\"preserve\"></t></r>", part);
    }

    [Fact]
    public void A_workbook_built_in_code_saves_its_rich_text()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].RichText = new RichText(new TextRun("Sales "), new TextRun("2026", Bold with { Color = CellColor.FromRgb(0xC00000) }));

        var saved = XlsxWriterTests.Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        Assert.Equal(sheet["A1"].RichText, XlsxWriterTests.Reload(saved)["S"]["A1"].RichText);
    }
}
