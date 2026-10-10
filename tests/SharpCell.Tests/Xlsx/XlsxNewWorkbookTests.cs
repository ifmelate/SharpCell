using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>A workbook built in code is saved as a new .xlsx file.</summary>
public class XlsxNewWorkbookTests
{
    private static Workbook Build()
    {
        var workbook = new Workbook();
        var sales = workbook.AddSheet("Sales");
        sales["A1"].Value = "Item";
        sales["B1"].Value = "Units";
        sales["C1"].Value = "Price";
        sales["A2"].Value = "Pen";
        sales["B2"].Value = 10;
        sales["C2"].Value = 1.5;
        sales["A3"].Value = "Book";
        sales["B3"].Value = 2;
        sales["C3"].Value = 12.25;
        sales.AddTable("Orders", "A1:C3");
        sales["E1"].Formula = "=SUMPRODUCT(Orders[Units],Orders[Price])";
        sales["E2"].Formula = "=SEQUENCE(3)";
        sales["F1"].Value = true;
        sales["F2"].Value = CellValue.Error(ErrorKind.NA);
        sales["F3"].Value = CellValue.DateTime(new DateTime(2026, 3, 15), workbook.DateSystem);
        sales["F3"].NumberFormat = "yyyy-mm-dd";
        sales["A1"].Style = CellStyle.Default with { Font = CellFont.Default with { Bold = true }, Fill = CellColor.FromRgb(0xD9E1F2) };
        sales.SetColumnWidth(1, 18);
        sales.SetRowHeight(1, 24);
        sales.Merge("G1:H1");
        sales.FrozenRows = 1;
        var summary = workbook.AddSheet("Summary");
        summary["A1"].Formula = "=Sales!E1*(1+Tax)";
        workbook.DefineName("Tax", "=0.2");
        workbook.Recalculate();
        return workbook;
    }

    [Fact]
    public void A_workbook_built_in_code_reads_back_the_same()
    {
        var workbook = Build();

        var saved = XlsxWriterTests.Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        var reloaded = XlsxWriterTests.Reload(saved);
        Assert.Equal(Structure.Of(workbook), Structure.Of(reloaded));
        Assert.Equal(Presentation.Of(workbook), Presentation.Of(reloaded));
        foreach (var sheet in workbook.Sheets)
        {
            Assert.Equal(Describe(sheet), Describe(reloaded[sheet.Name]));
        }

        reloaded.Recalculate();
        Assert.Equal(workbook["Summary"]["A1"].Value, reloaded["Summary"]["A1"].Value);
        Assert.Equal(3, reloaded["Sales"]["E4"].Value.AsNumber());
    }

    private static List<string> Describe(Worksheet sheet) =>
        [.. sheet.Cells.Select(c => $"{c.Address} {c.Formula} {c.Value} {c.NumberFormat} {c.Style}")];

    [Fact]
    public void A_workbook_without_sheets_cannot_be_saved()
    {
        Assert.Throws<InvalidOperationException>(() => XlsxWriterTests.Save(new Workbook()));
    }

    [Fact]
    public void Saving_twice_gives_the_same_bytes()
    {
        var workbook = Build();

        Assert.Equal(XlsxWriterTests.Save(workbook), XlsxWriterTests.Save(workbook));
    }

    [Fact]
    public void A_clone_of_a_new_workbook_saves_too()
    {
        var clone = Build().Clone();
        clone["Sales"]["B2"].Value = 20;
        clone.Recalculate();

        Assert.Equal(20, XlsxWriterTests.Reload(XlsxWriterTests.Save(clone))["Sales"]["B2"].Value.AsNumber());
    }
}
