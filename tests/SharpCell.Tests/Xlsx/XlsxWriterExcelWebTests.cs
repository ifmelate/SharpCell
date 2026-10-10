using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>
/// Files written by <see cref="XlsxWriter"/> that Excel for the web opened, recalculated and saved
/// (tests/corpus/excel-web/writer-*.xlsx): a workbook built in code, and Excel's templates with
/// rows inserted into every sheet, a style, a formula and a new sheet. Building them again here,
/// every formula and every value must be what Excel has, apart from volatile ones.
/// </summary>
public class XlsxWriterExcelWebTests
{
    public static TheoryData<string> Files() =>
        ["new", "crossword", "event_calendar", "gantt_project_tracker", "invoice", "travel_expenses_tracker", "weekly_timesheet", "wordle", "yearly_calendar"];

    [Theory]
    [MemberData(nameof(Files))]
    public void A_saved_workbook_has_the_formulas_and_values_Excel_has(string name)
    {
        var excel = XlsxReader.Load(CorpusFiles.PathOf($"excel-web/writer-{name}.xlsx"));

        // TODAY and NOW as when Excel calculated its copy: the templates' NOW cell says when that was.
        var now = excel.Sheets.SelectMany(s => s.Cells).FirstOrDefault(c => c.Formula == "=NOW()")?.Value.AsDateTime(excel.DateSystem);
        var clock = new FixedClock(now ?? new DateTime(2026, 10, 10, 23, 58, 0));
        var workbook = name == "new" ? BuildNew() : EditTemplate(CorpusFiles.PathOf($"excel/templates/{name}.xlsx"), clock);
        var ours = XlsxWriterTests.Reload(XlsxWriterTests.Save(workbook, new XlsxWriteOptions { KeepUncalculated = true }));

        Assert.Equal(ours.Sheets.Select(s => s.Name), excel.Sheets.Select(s => s.Name));
        Assert.Equal(ours.Tables.Select(t => $"{t.Name} {t.Range}"), excel.Tables.Select(t => $"{t.Name} {t.Range}"));
        foreach (var sheet in ours.Sheets)
        {
            var other = excel[sheet.Name];
            foreach (var (row, column) in sheet.Cells.Select(c => (c.Row, c.Column)).Union(other.Cells.Select(c => (c.Row, c.Column))))
            {
                var mine = sheet[row, column];
                var theirs = other[row, column];
                if (mine.Formula?.Contains("NOW(", StringComparison.Ordinal) == true)
                    continue;
                Assert.Equal((mine.ToString(), mine.Formula), (theirs.ToString(), theirs.Formula));
                if (mine.Value.Kind == CellValueKind.Number && theirs.Value.Kind == CellValueKind.Number)
                    Assert.Equal(theirs.Value.AsNumber(), mine.Value.AsNumber(), 9);
                else
                    Assert.Equal((mine.ToString(), theirs.Value), (theirs.ToString(), mine.Value));
            }
        }
    }

    // The workbook of writer-new.xlsx.
    private static Workbook BuildNew()
    {
        var workbook = new Workbook();
        var data = workbook.AddSheet("Data");
        string[] items = ["Pen", "Book", "Lamp", "Desk", "Chair"];
        double[] units = [10, 2, 4, 1, 6];
        double[] prices = [1.5, 12.25, 30, 199.99, 85];
        data["A1"].Value = "Item";
        data["B1"].Value = "Units";
        data["C1"].Value = "Price";
        data["D1"].Value = "Total";
        for (var i = 0; i < items.Length; i++)
        {
            data[i + 2, 1].Value = items[i];
            data[i + 2, 2].Value = units[i];
            data[i + 2, 3].Value = prices[i];
            data[i + 2, 3].NumberFormat = "#,##0.00";
        }

        data.AddTable("Orders", "A1:D6");
        for (var r = 2; r <= 6; r++)
            data[r, 4].Formula = "=[@Units]*[@Price]";
        var header = CellStyle.Default with
        {
            Font = CellFont.Default with { Bold = true, Color = CellColor.FromRgb(0xFFFFFF) },
            Fill = CellColor.FromRgb(0x1F4E79),
            HorizontalAlignment = CellHorizontalAlignment.Center,
        };
        for (var c = 1; c <= 4; c++)
            data[1, c].Style = header;
        data.SetColumnWidth(1, 16);
        data.FrozenRows = 1;
        var report = workbook.AddSheet("Report");
        string[] formulas =
        [
            "=SUM(Orders[Total])", "=XLOOKUP(\"Lamp\",Orders[Item],Orders[Total])", "=TEXTJOIN(\", \",TRUE,SORT(Orders[Item]))",
            "=LET(t,SUM(Orders[Total]),n,COUNTA(Orders[Item]),t/n)", "=LAMBDA(x,y,x*y+1)(3,4)", "=IFS(B1>500,\"big\",TRUE,\"small\")",
            "=ROWS(FILTER(Orders[Item],Orders[Units]>2))", "=CONCAT(\"n=\",COUNT(Orders[Units]))", "=MAXIFS(Orders[Price],Orders[Units],\">1\")",
            "=SUM(Data!B2:B6*Data!C2:C6)", "=STDEV.S(Orders[Price])", "=SWITCH(2,1,\"one\",2,\"two\")", "=SUM(F1#)", "=@Data!B2:B6",
            "=TaxRate*B1", "=TEXT(B1,\"#,##0.00\")",
        ];
        for (var i = 0; i < formulas.Length; i++)
        {
            report[i + 1, 1].Value = formulas[i].TrimStart('=');
            report[i + 1, 2].Formula = formulas[i];
        }

        report["F1"].Formula = "=SEQUENCE(4,1,10,10)";
        report["G1"].Formula = "=UNIQUE(Data!B2:B6>2)";
        workbook.DefineName("TaxRate", "=0.2");
        report.Merge("D1:E2");
        report["D1"].Value = "Merged";
        report.SetColumnWidth(1, 45);
        workbook.Recalculate();
        return workbook;
    }

    // An Excel template with two rows inserted at row 4 of every sheet, an italic note, and a new
    // sheet with a formula, a spill and a number format.
    private static Workbook EditTemplate(string path, TimeProvider clock)
    {
        var book = XlsxReader.Load(path);
        book.Clock = clock;
        book.Recalculate();
        var first = book.Sheets[0];
        foreach (var sheet in book.Sheets)
        {
            try
            {
                sheet.InsertRows(4, 2);
            }
            catch (Exception ex) when (ex is InvalidOperationException or NotSupportedException)
            {
                // Sheets whose arrays or tables would be split stay as they are.
            }
        }

        first["A4"].Value = "added by SharpCell";
        first["A4"].Style = first["A4"].Style with { Font = first["A4"].Style.Font with { Italic = true } };
        var added = book.AddSheet("SharpCell");
        added["A1"].Formula = $"=COUNTA('{first.Name}'!A1:Z200)";
        added["A2"].Formula = "=SEQUENCE(3)";
        added["B1"].Value = 1234.5;
        added["B1"].NumberFormat = "#,##0.00 \"EUR\"";
        book.Recalculate();
        return book;
    }

    private sealed class FixedClock(DateTime local) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => new(local, TimeSpan.Zero);

        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }
}
