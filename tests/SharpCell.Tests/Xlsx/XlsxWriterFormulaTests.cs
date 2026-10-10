using System.Linq;
using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

/// <summary>New, changed and removed formulas are saved, in the syntax files use.</summary>
public class XlsxWriterFormulaTests
{
    private static TestXlsx Book() => new TestXlsx().Sheet("S",
        "<row r=\"1\"><c r=\"A1\"><v>2</v></c><c r=\"B1\"><v>7</v></c></row>"
        + "<row r=\"2\"><c r=\"A2\"><f>A1*3</f><v>6</v></c></row>");

    private static (byte[] Original, byte[] Saved, Workbook Reloaded) Edit(TestXlsx file, Action<Worksheet> change)
    {
        var original = file.Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        change(workbook["S"]);
        workbook.Recalculate();
        var saved = Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        return (original, saved, Reload(saved));
    }

    [Fact]
    public void A_constant_replaced_by_a_formula_is_saved_with_its_value()
    {
        var (_, saved, reloaded) = Edit(Book(), s => s["B1"].Formula = "=A1*2");

        Assert.Contains("<c r=\"B1\"><f>A1*2</f><v>4</v></c>", Part(saved, "xl/worksheets/sheet1.xml"));
        Assert.Equal("=A1*2", reloaded["S"]["B1"].Formula);
        Assert.Equal(4, reloaded["S"]["B1"].Value.AsNumber());
    }

    [Fact]
    public void A_changed_formula_is_saved()
    {
        var (_, _, reloaded) = Edit(Book(), s => s["A2"].Formula = "=A1+B1");

        Assert.Equal("=A1+B1", reloaded["S"]["A2"].Formula);
        Assert.Equal(9, reloaded["S"]["A2"].Value.AsNumber());
    }

    [Fact]
    public void A_formula_replaced_by_its_own_value_loses_the_formula()
    {
        var (_, _, reloaded) = Edit(Book(), s => s["A2"].Value = 6);

        Assert.Null(reloaded["S"]["A2"].Formula);
        Assert.Equal(6, reloaded["S"]["A2"].Value.AsNumber());
    }

    [Fact]
    public void A_removed_formula_removes_the_cell()
    {
        var (_, saved, reloaded) = Edit(Book(), s => s["A2"].Formula = null);

        Assert.DoesNotContain("A2", Part(saved, "xl/worksheets/sheet1.xml"));
        Assert.Equal(CellValue.Empty, reloaded["S"]["A2"].Value);
    }

    [Fact]
    public void A_formula_in_a_new_cell_and_a_new_row_is_saved()
    {
        var (_, _, reloaded) = Edit(Book(), s =>
        {
            s["C1"].Formula = "=SUM(A1:B1)";
            s["D5"].Formula = "=\"n=\"&C1";
        });

        Assert.Equal(("=SUM(A1:B1)", 9.0), (reloaded["S"]["C1"].Formula, reloaded["S"]["C1"].Value.AsNumber()));
        Assert.Equal(("=\"n=\"&C1", "n=9"), (reloaded["S"]["D5"].Formula, reloaded["S"]["D5"].Value.AsText()));
    }

    [Fact]
    public void Newer_functions_are_saved_with_their_prefixes()
    {
        var (_, saved, reloaded) = Edit(Book(), s => s["C1"].Formula = "=LET(x,A1,XLOOKUP(x,A1:A2,A1:A2))");

        Assert.Contains(">_xlfn.LET(_xlpm.x,A1,_xlfn.XLOOKUP(_xlpm.x,A1:A2,A1:A2))</f>", Part(saved, "xl/worksheets/sheet1.xml"));
        Assert.Equal(2, reloaded["S"]["C1"].Value.AsNumber());
    }

    [Fact]
    public void A_spilling_formula_is_saved_as_a_dynamic_array_with_metadata()
    {
        var (_, saved, reloaded) = Edit(Book(), s => s["C1"].Formula = "=SEQUENCE(3)");

        var sheet = Part(saved, "xl/worksheets/sheet1.xml");
        Assert.Contains("<c r=\"C1\" cm=\"1\"><f t=\"array\" ref=\"C1:C3\">_xlfn.SEQUENCE(3)</f><v>1</v></c>", sheet);
        Assert.Contains("XLDAPR", Part(saved, "xl/metadata.xml"));
        Assert.Contains("sheetMetadata", Part(saved, "xl/_rels/workbook.xml.rels"));
        Assert.Contains("/xl/metadata.xml", Part(saved, "[Content_Types].xml"));
        Assert.Equal(3, reloaded["S"]["C3"].Value.AsNumber());
        reloaded.Recalculate();
        Assert.Equal(3, reloaded["S"]["C3"].Value.AsNumber());   // read back as a dynamic array
    }

    [Fact]
    public void A_file_with_dynamic_array_metadata_has_its_entry_reused()
    {
        var file = Book();
        file.Metadata = TestXlsx.DynamicArrayMetadata;
        var (_, saved, _) = Edit(file, s => s["C1"].Formula = "=SEQUENCE(2)");

        Assert.Contains("<c r=\"C1\" cm=\"1\">", Part(saved, "xl/worksheets/sheet1.xml"));
        Assert.Equal(TestXlsx.DynamicArrayMetadata, Part(saved, "xl/metadata.xml"));
    }

    [Fact]
    public void A_formula_reading_a_range_where_one_value_is_expected_is_saved_as_a_dynamic_array()
    {
        var (_, saved, reloaded) = Edit(Book(), s => s["C2"].Formula = "=SUM(A1:A2*2)");

        Assert.Contains("<c r=\"C2\" cm=\"1\"><f t=\"array\" ref=\"C2\">SUM(A1:A2*2)</f><v>16</v></c>", Part(saved, "xl/worksheets/sheet1.xml"));
        reloaded.Recalculate();
        Assert.Equal(16, reloaded["S"]["C2"].Value.AsNumber());
    }

    [Fact]
    public void A_shared_formula_group_whose_master_changes_is_saved_in_full()
    {
        var file = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><f t=\"shared\" ref=\"B1:B3\" si=\"0\">A1*10</f><v>10</v></c></row>"
            + "<row r=\"2\"><c r=\"A2\"><v>2</v></c><c r=\"B2\"><f t=\"shared\" si=\"0\"/><v>20</v></c></row>"
            + "<row r=\"3\"><c r=\"A3\"><v>3</v></c><c r=\"B3\"><f t=\"shared\" si=\"0\"/><v>30</v></c></row>");
        var (_, saved, reloaded) = Edit(file, s => s["B1"].Formula = "=A1*100");

        Assert.DoesNotContain("shared", Part(saved, "xl/worksheets/sheet1.xml"));
        Assert.Equal(["=A1*100", "=A2*10", "=A3*10"], new[] { "B1", "B2", "B3" }.Select(a => reloaded["S"][a].Formula));
        Assert.Equal(30, reloaded["S"]["B3"].Value.AsNumber());
    }

    [Fact]
    public void Changing_formulas_drops_the_calc_chain()
    {
        var file = Book();
        file.CalcChain = ["A2"];
        var (_, saved, _) = Edit(file, s => s["B1"].Formula = "=A1");

        Assert.Null(PartOrNull(saved, "xl/calcChain.xml"));
        Assert.DoesNotContain("calcChain", Part(saved, "[Content_Types].xml"));
        Assert.DoesNotContain("calcChain", Part(saved, "xl/_rels/workbook.xml.rels"));
    }

    [Fact]
    public void Changing_only_values_keeps_the_calc_chain()
    {
        var file = Book();
        file.CalcChain = ["A2"];
        var (_, saved, _) = Edit(file, s => s["A1"].Value = 5);

        Assert.NotNull(PartOrNull(saved, "xl/calcChain.xml"));
    }

    private static string? PartOrNull(byte[] file, string name)
    {
        using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(file));
        return zip.GetEntry(name) is null ? null : Part(file, name);
    }
}
