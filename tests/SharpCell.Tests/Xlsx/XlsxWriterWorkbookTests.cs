using System.IO;
using System.Linq;
using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

/// <summary>New sheets, defined names, tables and workbook settings are saved.</summary>
public class XlsxWriterWorkbookTests
{
    private static TestXlsx Book() => new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>2</v></c></row><row r=\"2\"><c r=\"A2\"><f>A1*3</f><v>6</v></c></row>")
        .ChartSheet("Chart")
        .Sheet("T", "<row r=\"1\"><c r=\"A1\"><v>5</v></c></row>")
        .Name("_xlnm.Print_Area", "S!$A$1:$B$2", " localSheetId=\"0\"")
        .Name("Rate", "0.1");

    private static (byte[] Saved, Workbook Model, Workbook Reloaded) Edit(Action<Workbook> change, TestXlsx? file = null)
    {
        var original = (file ?? Book()).Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        change(workbook);
        workbook.Recalculate();
        var saved = Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        return (saved, workbook, Reload(saved));
    }

    [Theory]
    [MemberData(nameof(XlsxSourceTests.Changes), MemberType = typeof(XlsxSourceTests))]
    public void A_structure_change_is_saved(string change)
    {
        var (_, model, reloaded) = Edit(workbook => XlsxSourceTests.Change(workbook, change), XlsxSourceTests.Template());

        Assert.Equal(Structure.Of(model), Structure.Of(reloaded));
    }

    [Fact]
    public void A_new_sheet_is_saved_with_its_cells()
    {
        var (saved, _, reloaded) = Edit(workbook =>
        {
            var sheet = workbook.AddSheet("New Sheet");
            sheet["A1"].Value = "text";
            sheet["B2"].Formula = "=S!A1*10";
            sheet["B2"].Style = CellStyle.Default with { Font = CellFont.Default with { Bold = true } };
            workbook["S"]["B1"].Formula = "='New Sheet'!B2+1";
        });

        Assert.Equal(["S", "T", "New Sheet"], reloaded.Sheets.Select(s => s.Name));
        var sheet = reloaded["New Sheet"];
        Assert.Equal(("text", "=S!A1*10", 20.0), (sheet["A1"].Value.AsText(), sheet["B2"].Formula, sheet["B2"].Value.AsNumber()));
        Assert.True(sheet["B2"].Style.Font.Bold);
        Assert.Equal(21, reloaded["S"]["B1"].Value.AsNumber());
        Assert.Contains("sheetId=\"4\"", Part(saved, "xl/workbook.xml"));
    }

    [Fact]
    public void Two_new_sheets_get_their_own_parts()
    {
        var (_, _, reloaded) = Edit(workbook =>
        {
            workbook.AddSheet("A")["A1"].Value = 1;
            workbook.AddSheet("B")["A1"].Value = 2;
        });

        Assert.Equal((1.0, 2.0), (reloaded["A"]["A1"].Value.AsNumber(), reloaded["B"]["A1"].Value.AsNumber()));
    }

    [Fact]
    public void Defined_names_are_saved_and_names_SharpCell_does_not_read_are_kept()
    {
        var (saved, _, reloaded) = Edit(workbook =>
        {
            workbook.DefineName("Rate", "=0.2");
            workbook.DefineName("Total", "=SUM(S!$A$1:$A$2)");
            workbook.DefineName("Local", "=T!$A$1", workbook["T"]);
            workbook.DefineName("Pick", "=XLOOKUP(1,S!$A:$A,S!$A:$A)");
        });

        Assert.Equal(["Rate =0.2 (workbook)", "Total =SUM(S!$A$1:$A$2) (workbook)", "Local =T!$A$1 T", "Pick =XLOOKUP(1,S!$A:$A,S!$A:$A) (workbook)"],
            reloaded.DefinedNames.Select(n => $"{n.Name} {n.Formula} {n.Scope?.Name ?? "(workbook)"}"));
        var workbookXml = Part(saved, "xl/workbook.xml");
        Assert.Contains("<definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">S!$A$1:$B$2</definedName>", workbookXml);
        Assert.Contains("<definedName name=\"Local\" localSheetId=\"2\">T!$A$1</definedName>", workbookXml);
        Assert.Contains("_xlfn.XLOOKUP", workbookXml);
    }

    [Fact]
    public void A_new_table_is_saved_with_its_columns()
    {
        var (saved, _, reloaded) = Edit(workbook =>
        {
            var sheet = workbook["T"];
            sheet["C1"].Value = "Units";
            sheet["C2"].Value = 3;
            sheet["D2"].Value = 4;
            sheet.AddTable("Sales", "C1:D3");
            sheet["E2"].Formula = "=SUM(Sales[Units])";
        });

        var table = Assert.Single(reloaded.Tables);
        Assert.Equal(("Sales", "C1:D3", "T"), (table.Name, table.Range, table.Worksheet.Name));
        Assert.Equal(["Units", "Column2"], table.Columns);
        Assert.Equal("Column2", reloaded["T"]["D1"].Value.AsText());   // header cells hold the column names, as Excel needs
        Assert.Equal(3, reloaded["T"]["E2"].Value.AsNumber());
        Assert.Matches("<tableParts [^>]*count=\"1\"", Part(saved, "xl/worksheets/sheet3.xml"));
    }

    [Fact]
    public void Date_system_and_iteration_settings_are_saved()
    {
        var (_, _, reloaded) = Edit(workbook =>
        {
            workbook.DateSystem = DateSystem.Date1904;
            workbook.Iteration = new IterationSettings(Enabled: true, MaxIterations: 50, MaxChange: 0.01);
        });

        Assert.Equal(DateSystem.Date1904, reloaded.DateSystem);
        Assert.Equal(new IterationSettings(true, 50, 0.01), reloaded.Iteration);
    }

    // What Excel for the web wrote for the password "sharpcell" on a sheet; a workbook's hash is made the same way.
    private const string ProtectionElement =
        """<workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="ZsXfA+OtvfMzrF4AroCZoBBxx+a92KoNFW4/RiaI5gOZ34t5Xky94AX76colv22hxkIqKFlpW8FJHO0Wyud2mQ==" workbookSaltValue="xCW5ZRsTr/6hfBxYoNlpbA==" workbookSpinCount="100000" lockStructure="1"/>""";

    private static string? ProtectionAttributes(byte[] saved)
    {
        var element = System.Xml.Linq.XDocument.Parse(Part(saved, "xl/workbook.xml")).Root!.Elements().SingleOrDefault(e => e.Name.LocalName == "workbookProtection");
        return element is null ? null : string.Join(" ", element.Attributes().Select(a => $"{a.Name.LocalName}={a.Value}"));
    }

    [Fact]
    public void A_protected_workbook_is_saved_with_the_hash_and_read_back()
    {
        var (saved, _, reloaded) = Edit(workbook => workbook.Protect("pw"));

        var attributes = ProtectionAttributes(saved)!.Split(' ');
        Assert.Equal(["workbookAlgorithmName=SHA-512", "workbookSpinCount=100000", "lockStructure=1"],
            attributes.Where(a => !a.StartsWith("workbookHashValue") && !a.StartsWith("workbookSaltValue")));
        Assert.Contains(attributes, a => a.StartsWith("workbookHashValue="));
        Assert.Contains(attributes, a => a.StartsWith("workbookSaltValue="));
        Assert.DoesNotContain("lockWindows", ProtectionAttributes(saved));
        Assert.True(reloaded.IsStructureProtected);
        Assert.False(reloaded.IsWindowsProtected);
        Assert.True(reloaded.ProtectionPassword!.Verify("pw"));
        Assert.False(reloaded.ProtectionPassword.Verify("PW"));
    }

    [Fact]
    public void Locked_windows_are_saved()
    {
        var (saved, _, reloaded) = Edit(workbook => workbook.Protect(structure: false, windows: true));

        Assert.Equal("lockWindows=1", ProtectionAttributes(saved));
        Assert.False(reloaded.IsStructureProtected);
        Assert.True(reloaded.IsWindowsProtected);
        Assert.Null(reloaded.ProtectionPassword);
    }

    [Fact]
    public void The_protection_element_is_placed_where_the_schema_wants_it()
    {
        var (saved, _, _) = Edit(workbook => workbook.Protect("pw"), new TestXlsx { WorkbookExtra = "<workbookPr/>" }.Sheet("S", ""));

        var names = System.Xml.Linq.XDocument.Parse(Part(saved, "xl/workbook.xml")).Root!.Elements().Select(e => e.Name.LocalName);
        Assert.Equal(["workbookPr", "workbookProtection", "sheets"], names);
    }

    [Fact]
    public void A_protection_the_file_has_is_read()
    {
        var file = new TestXlsx { WorkbookExtra = ProtectionElement }.Sheet("S", "");
        var workbook = XlsxReader.Load(new MemoryStream(file.Build().ToArray()));

        Assert.True(workbook.IsStructureProtected);
        Assert.False(workbook.IsWindowsProtected);
        Assert.True(workbook.ProtectionPassword!.Verify("sharpcell"));
    }

    [Fact]
    public void A_protected_workbook_saved_without_changes_is_the_source_bytes()
    {
        var original = new TestXlsx { WorkbookExtra = ProtectionElement }.Sheet("S", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>").Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();

        Assert.Equal(original, Save(workbook));
    }

    [Fact]
    public void Unprotecting_removes_the_element()
    {
        var file = new TestXlsx { WorkbookExtra = ProtectionElement }.Sheet("S", "");
        var (saved, _, reloaded) = Edit(workbook => workbook.Unprotect(), file);

        Assert.Null(ProtectionAttributes(saved));
        Assert.False(reloaded.IsStructureProtected);
        Assert.Null(reloaded.ProtectionPassword);
    }

    [Fact]
    public void Other_attributes_of_the_element_are_kept_when_it_is_rewritten()
    {
        var file = new TestXlsx { WorkbookExtra = """<workbookProtection lockStructure="1" lockRevision="1" revisionsAlgorithmName="SHA-512" revisionsHashValue="AAAA"/>""" }.Sheet("S", "");
        var (saved, _, reloaded) = Edit(workbook => workbook.Protect("pw", structure: false, windows: true), file);

        var attributes = ProtectionAttributes(saved)!;
        Assert.Contains("lockRevision=1", attributes);
        Assert.Contains("revisionsHashValue=AAAA", attributes);
        Assert.Contains("lockWindows=1", attributes);
        Assert.DoesNotContain("lockStructure", attributes);
        Assert.True(reloaded.ProtectionPassword!.Verify("pw"));

        var (unprotected, _, _) = Edit(workbook => workbook.Unprotect(), file);
        Assert.Equal("lockRevision=1 revisionsAlgorithmName=SHA-512 revisionsHashValue=AAAA", ProtectionAttributes(unprotected));
    }

    [Fact]
    public void A_legacy_password_is_written_back_unchanged()
    {
        var file = new TestXlsx { WorkbookExtra = """<workbookProtection workbookPassword="CC1A" lockStructure="1"/>""" }.Sheet("S", "");
        var (saved, _, _) = Edit(workbook => workbook.AddSheet("New"), file);

        Assert.Equal("workbookPassword=CC1A lockStructure=1", ProtectionAttributes(saved));
    }

    [Fact]
    public void A_cloned_protected_workbook_saves_its_protection()
    {
        var workbook = new Workbook();
        workbook.AddSheet("S");
        workbook.Protect("pw", windows: true);

        var reloaded = Reload(Save(workbook.Clone()));

        Assert.True(reloaded.IsStructureProtected);
        Assert.True(reloaded.IsWindowsProtected);
        Assert.True(reloaded.ProtectionPassword!.Verify("pw"));
    }
}
