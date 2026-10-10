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
}
