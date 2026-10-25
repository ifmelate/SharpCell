using System.Data;
using System.Globalization;

namespace SharpCell.Tests.Docs;

/// <summary>The code shown in docs/guides/sheets.md and data.md. See <see cref="GuideExamples"/> for how snippets work.</summary>
public class SheetGuideExamples
{
    private static (Workbook Workbook, Worksheet Sheet) Sales()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sales");
        sheet["A1"].Value = "Region";
        sheet["B1"].Value = "Units";
        string[] regions = ["North", "South", "East", "West"];
        double[] units = [120, 340, 75, 410];
        for (var i = 0; i < regions.Length; i++)
        {
            sheet[i + 2, 1].Value = regions[i];
            sheet[i + 2, 2].Value = units[i];
        }

        workbook.Recalculate();
        return (workbook, sheet);
    }

    [Fact]
    public void Manage_sheets()
    {
        var workbook = new Workbook();
        var summary = workbook.AddSheet("Summary");
        var data = workbook.AddSheet("Data");
        var draft = workbook.AddSheet("Draft");
        data["A1"].Value = 5;
        summary["A1"].Formula = "=Data!A1*2";
        draft["A1"].Formula = "=Data!A1";

        // snippet: sheets-manage
        data.Name = "Data 2025";             // Summary!A1 is now ='Data 2025'!A1*2
        workbook.MoveSheet(data, 0);         // Data 2025 comes first
        workbook.RemoveSheet(draft);         // formulas reading Draft get #REF!
        data.Visibility = SheetVisibility.Hidden;
        summary.TabColor = CellColor.FromRgb(0x2E75B6);
        workbook.Recalculate();
        // end-snippet

        Assert.Equal(("='Data 2025'!A1*2", 10.0), (summary["A1"].Formula, summary["A1"].Value.AsNumber()));
        Assert.Equal([data, summary], workbook.Sheets);
        Assert.True(draft.IsRemoved);
    }

    [Fact]
    public void Filter_rows()
    {
        var (_, sheet) = Sales();

        // snippet: sheets-filter
        AutoFilter filter = sheet.SetAutoFilter("A1:B5");
        filter.Filter(2, ColumnFilter.Custom(FilterOperator.GreaterThan, "100"));   // column B
        filter.Filter(1, ColumnFilter.OneOf("North", "West"));                      // column A
        bool hidden = sheet.IsRowHidden(3);   // true: South is not North or West

        sheet["B2"].Value = 90;
        filter.Apply();                       // Excel's Reapply: North is hidden now
        // end-snippet

        Assert.True(hidden);
        Assert.Equal([false, true, true, true, false], Enumerable.Range(1, 5).Select(sheet.IsRowHidden));
    }

    [Fact]
    public void Sort_rows()
    {
        var (_, sheet) = Sales();

        // snippet: sheets-sort
        sheet.Range("A2:B5").Sort(new SortKey(2, Descending: true));   // by units, highest first
        // end-snippet

        Assert.Equal(["West", "South", "North", "East"], Enumerable.Range(2, 4).Select(row => sheet[row, 1].Text));
    }

    [Fact]
    public void Data_out()
    {
        var (_, sheet) = Sales();
        var writer = new StringWriter(CultureInfo.InvariantCulture);

        // snippet: data-out
        DataTable table = sheet.Range("A1:B5").ToDataTable();   // Region (string), Units (double)

        List<(string Region, double Units)> rows = sheet.Range("A1:B5")
            .ReadRows(row => (row["Region"].Text, row["Units"].Value.AsNumber()));

        string csv = sheet.Range("A1:B5").ToCsv();               // "Region,Units\r\nNorth,120\r\n..."
        sheet.Range("A1:B5").WriteCsv(writer);                   // to a file or a stream
        // end-snippet

        Assert.Equal(typeof(double), table.Columns["Units"]!.DataType);
        Assert.Equal(("West", 410.0), rows[3]);
        Assert.StartsWith("Region,Units\r\nNorth,120\r\n", csv);
        Assert.Equal(csv, writer.ToString());
    }

    [Fact]
    public void Data_in()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Orders");
        var orders = new[] { (Id: 1, Total: 99.5m, Placed: new DateTime(2026, 3, 5)) };

        // snippet: data-in
        sheet.LoadRows("A1", orders,
            ("Id", o => o.Id), ("Total", o => o.Total), ("Placed", o => o.Placed));

        sheet.LoadCsv("E1", new StringReader("Code,Share\nX1,12%\n"));   // 12% is the number 0.12
        // end-snippet

        Assert.Equal(99.5, sheet["B2"].Value.AsNumber());
        Assert.Equal(0.12, sheet["F2"].Value.AsNumber());
    }
}
