using System.Linq;
using System.Xml.Linq;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>
/// Sheet protection is read from and saved in the sheet's &lt;sheetProtection&gt; element. Its flags state
/// what is protected, the opposite of <see cref="SheetProtection"/>, which states what is allowed.
/// </summary>
public class XlsxWriterProtectionTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    // What Excel for the web wrote for the password "sharpcell" (tests/corpus/excel-web/protection.xlsx).
    private const string ExcelHash = "ZsXfA+OtvfMzrF4AroCZoBBxx+a92KoNFW4/RiaI5gOZ34t5Xky94AX76colv22hxkIqKFlpW8FJHO0Wyud2mQ==";
    private const string ExcelSalt = "xCW5ZRsTr/6hfBxYoNlpbA==";

    private const string ExcelProtection =
        $"""<sheetProtection algorithmName="SHA-512" hashValue="{ExcelHash}" saltValue="{ExcelSalt}" spinCount="100000" sheet="1" objects="1" scenarios="1" formatCells="0" sort="0"/>""";

    private static TestXlsx Sheet(string protection) => new TestXlsx().RawSheet("S", $"""
        <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
        <sheetCalcPr fullCalcOnLoad="1"/>
        {protection}
        <mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells>
        <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        """);

    private static Workbook Load(byte[] file) => XlsxReader.Load(new MemoryStream(file));

    // The attributes of the saved sheet's sheetProtection element, in order; null when it has none.
    private static string? Attributes(byte[] saved, string part = Sheet1)
    {
        var element = XDocument.Parse(XlsxWriterTests.Part(saved, part)).Root!.Elements().SingleOrDefault(e => e.Name.LocalName == "sheetProtection");
        return element is null ? null : string.Join(" ", element.Attributes().Select(a => $"{a.Name.LocalName}={a.Value}"));
    }

    private static string[] ElementNames(byte[] saved) =>
        [.. XDocument.Parse(XlsxWriterTests.Part(saved, Sheet1)).Root!.Elements().Select(e => e.Name.LocalName)];

    [Fact]
    public void Excel_hashes_a_password_as_SharpCell_checks_it()
    {
        var workbook = XlsxReader.Load(CorpusFiles.PathOf("excel-web/protection.xlsx"));
        var sheet = workbook.Sheets[0];

        Assert.Equal(new SheetProtection(), sheet.Protection);
        Assert.Equal("SHA-512", sheet.ProtectionPassword!.AlgorithmName);
        Assert.True(sheet.ProtectionPassword.Verify("sharpcell"));
        Assert.False(sheet.ProtectionPassword.Verify("Sharpcell"));
    }

    [Fact]
    public void Flags_the_file_turns_off_are_allowed_actions()
    {
        var sheet = Load(Sheet(ExcelProtection).Build().ToArray())["S"];

        Assert.Equal(new SheetProtection { FormatCells = true, Sort = true }, sheet.Protection);
        Assert.True(sheet.ProtectionPassword!.Verify("sharpcell"));
    }

    [Fact]
    public void A_sheet_without_the_sheet_flag_is_not_protected()
    {
        var sheet = Load(Sheet("<sheetProtection formatCells=\"0\"/>").Build().ToArray())["S"];

        Assert.Null(sheet.Protection);
        Assert.Null(Load(Sheet("").Build().ToArray())["S"].Protection);
    }

    [Fact]
    public void A_protected_workbook_saved_without_changes_is_the_source_bytes()
    {
        var original = Sheet(ExcelProtection).Build().ToArray();
        var workbook = Load(original);
        workbook.Recalculate();

        Assert.Equal(original, XlsxWriterTests.Save(workbook));
    }

    [Fact]
    public void The_hash_is_written_back_unchanged_after_a_value_change()
    {
        var original = Sheet(ExcelProtection).Build().ToArray();
        var workbook = Load(original);
        workbook["S"]["A1"].Value = 2;
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        Assert.Equal(Attributes(original), Attributes(saved));
        Assert.True(Load(saved)["S"].ProtectionPassword!.Verify("sharpcell"));
    }

    [Fact]
    public void Allowed_actions_are_written_as_Excel_inverts_them()
    {
        var workbook = new Workbook();
        workbook.AddSheet("Defaults").Protect("secret");
        workbook.AddSheet("Some").Protect(allow: new SheetProtection { FormatCells = true, InsertRows = true, Sort = true, EditObjects = true, SelectLockedCells = false });
        workbook.AddSheet("Plain");
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        var defaults = XDocument.Parse(XlsxWriterTests.Part(saved, Sheet1)).Root!.Elements().Single(e => e.Name.LocalName == "sheetProtection");
        Assert.Equal(["algorithmName", "hashValue", "saltValue", "spinCount", "sheet", "objects", "scenarios"], defaults.Attributes().Select(a => a.Name.LocalName));
        Assert.Equal("SHA-512", (string?)defaults.Attribute("algorithmName"));
        Assert.Equal("100000", (string?)defaults.Attribute("spinCount"));
        Assert.Equal(["1", "1", "1"], new[] { "sheet", "objects", "scenarios" }.Select(name => (string?)defaults.Attribute(name)));
        Assert.Equal("sheet=1 scenarios=1 formatCells=0 insertRows=0 selectLockedCells=1 sort=0", Attributes(saved, "xl/worksheets/sheet2.xml"));
        Assert.Null(Attributes(saved, "xl/worksheets/sheet3.xml"));

        var reloaded = Load(saved);
        Assert.Equal(new SheetProtection(), reloaded["Defaults"].Protection);
        Assert.True(reloaded["Defaults"].ProtectionPassword!.Verify("secret"));
        Assert.Equal(workbook["Some"].Protection, reloaded["Some"].Protection);
        Assert.Null(reloaded["Some"].ProtectionPassword);
        Assert.Null(reloaded["Plain"].Protection);
    }

    public static TheoryData<string> Actions => [.. typeof(SheetProtection).GetProperties().Where(p => p.PropertyType == typeof(bool)).Select(p => p.Name)];

    [Theory]
    [MemberData(nameof(Actions))]
    public void Each_allowed_action_reads_back(string action)
    {
        var allow = new SheetProtection();
        var property = typeof(SheetProtection).GetProperty(action)!;
        property.SetValue(allow, !(bool)property.GetValue(allow)!);
        var workbook = new Workbook();
        workbook.AddSheet("S").Protect(allow: allow);

        var saved = XlsxWriterTests.Save(workbook);

        Assert.Equal(allow, Load(saved)["S"].Protection);
        Assert.NotEqual(new SheetProtection(), allow);
    }

    [Fact]
    public void A_protection_added_to_a_file_lands_where_the_schema_puts_it()
    {
        var original = Sheet("").Build().ToArray();
        var workbook = Load(original);
        workbook["S"].Protect("secret", new SheetProtection { Sort = true });
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        Assert.Equal(["sheetData", "sheetCalcPr", "sheetProtection", "mergeCells", "pageMargins"], ElementNames(saved));
        Assert.True(Load(saved)["S"].ProtectionPassword!.Verify("secret"));
    }

    [Fact]
    public void Protecting_again_replaces_the_element()
    {
        var original = Sheet(ExcelProtection).Build().ToArray();
        var workbook = Load(original);
        workbook["S"].Protect("other");
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        Assert.Equal(["sheetData", "sheetCalcPr", "sheetProtection", "mergeCells", "pageMargins"], ElementNames(saved));
        var reloaded = Load(saved)["S"];
        Assert.Equal(new SheetProtection(), reloaded.Protection);
        Assert.True(reloaded.ProtectionPassword!.Verify("other"));
        Assert.False(reloaded.ProtectionPassword.Verify("sharpcell"));
    }

    [Fact]
    public void Unprotecting_removes_the_element()
    {
        var original = Sheet(ExcelProtection).Build().ToArray();
        var workbook = Load(original);
        workbook["S"].Unprotect();
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        Assert.Equal(["sheetData", "sheetCalcPr", "mergeCells", "pageMargins"], ElementNames(saved));
        Assert.Null(Load(saved)["S"].Protection);
    }

    [Fact]
    public void An_element_without_the_sheet_flag_is_replaced_when_the_sheet_is_protected()
    {
        var original = Sheet("<sheetProtection formatCells=\"0\"/>").Build().ToArray();
        var workbook = Load(original);
        workbook.Recalculate();
        Assert.Equal(original, XlsxWriterTests.Save(workbook));

        workbook["S"].Protect();
        var saved = XlsxWriterTests.Save(workbook);

        Assert.Equal("sheet=1 objects=1 scenarios=1", Attributes(saved));
    }

    [Fact]
    public void A_legacy_password_is_kept_as_read()
    {
        var original = Sheet("<sheetProtection password=\"CC3D\" sheet=\"1\" objects=\"1\" scenarios=\"1\"/>").Build().ToArray();
        var workbook = Load(original);
        var sheet = workbook["S"];
        Assert.Equal(new SheetProtection(), sheet.Protection);
        Assert.Equal("CC3D", sheet.ProtectionPassword!.LegacyHash);
        Assert.False(sheet.ProtectionPassword.Verify("secret"));

        // Only the flags change (there is no public way to do that without a new password): the element is
        // written again with the password the file had.
        sheet.SetProtection(new SheetProtection { Sort = true }, sheet.ProtectionPassword);
        workbook.Recalculate();
        var saved = XlsxWriterTests.Save(workbook);

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        Assert.Equal("password=CC3D sheet=1 objects=1 scenarios=1 sort=0", Attributes(saved));
    }

    [Fact]
    public void A_hash_read_from_a_file_is_written_back_when_the_element_is_written_again()
    {
        var original = Sheet(ExcelProtection).Build().ToArray();
        var workbook = Load(original);
        var sheet = workbook["S"];
        sheet.SetProtection(new SheetProtection(), sheet.ProtectionPassword);
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        Assert.Equal($"algorithmName=SHA-512 hashValue={ExcelHash} saltValue={ExcelSalt} spinCount=100000 sheet=1 objects=1 scenarios=1", Attributes(saved));
    }
}
