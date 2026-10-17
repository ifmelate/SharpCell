using SharpCell.Tests.Xlsx;

namespace SharpCell.Tests.Model;

/// <summary>Sorting a range puts its rows (or columns) in order, as Excel's Data &gt; Sort does.</summary>
public class SortTests
{
    private static Worksheet Sheet(out Workbook workbook)
    {
        workbook = new Workbook();
        return workbook.AddSheet("S");
    }

    private static string Column(Worksheet sheet, int column, int first, int last) =>
        string.Join(",", Enumerable.Range(first, last - first + 1).Select(row => sheet[row, column].Value.Kind == CellValueKind.Empty ? "_" : sheet[row, column].Text));

    [Fact]
    public void Ascending_puts_numbers_text_logicals_errors_then_blanks()
    {
        var s = Sheet(out var w);
        s["A1"].Value = "b";
        s["A2"].Value = true;
        s["A3"].Formula = "=NA()";
        s["A4"].Value = 10;
        s["A6"].Value = "A";
        s["A7"].Value = false;
        s["A8"].Value = -1;
        s["A9"].Formula = "=1/0";
        w.Recalculate();

        s.Range("A1:A9").Sort();
        w.Recalculate();

        Assert.Equal("-1,10,A,b,FALSE,TRUE,#N/A,#DIV/0!,_", Column(s, 1, 1, 9));
    }

    [Fact]
    public void Descending_reverses_the_order_but_keeps_blanks_last()
    {
        var s = Sheet(out var w);
        s["A1"].Value = 1;
        s["A3"].Value = "x";
        s["A4"].Value = 3;
        s["A5"].Value = true;
        w.Recalculate();

        s.Range("A1:A5").Sort(new SortKey(1, Descending: true));

        Assert.Equal("TRUE,x,3,1,_", Column(s, 1, 1, 5));
    }

    [Fact]
    public void Rows_move_whole_ordered_by_several_keys_and_ties_keep_their_order()
    {
        var s = Sheet(out var w);
        string[] regions = ["East", "West", "East", "West", "East"];
        double[] units = [5, 3, 5, 9, 1];
        for (var i = 0; i < 5; i++)
        {
            s[i + 1, 1].Value = regions[i];
            s[i + 1, 2].Value = units[i];
            s[i + 1, 3].Value = i + 1;   // the original row
        }

        w.Recalculate();

        s.Range("A1:C5").Sort(new SortKey(1), new SortKey(2, Descending: true));

        Assert.Equal("East,East,East,West,West", Column(s, 1, 1, 5));
        Assert.Equal("5,5,1,9,3", Column(s, 2, 1, 5));
        Assert.Equal("1,3,5,4,2", Column(s, 3, 1, 5));
    }

    [Fact]
    public void Text_ignores_case_unless_asked_and_then_puts_lower_case_first()
    {
        var s = Sheet(out var w);
        s["A1"].Value = "B";
        s["A2"].Value = "a";
        s["A3"].Value = "b";
        s["A4"].Value = "A";
        w.Recalculate();

        s.Range("A1:A4").Sort();
        Assert.Equal("a,A,B,b", Column(s, 1, 1, 4));

        s.Range("A1:A4").Sort([new SortKey(1)], matchCase: true);
        Assert.Equal("a,A,b,B", Column(s, 1, 1, 4));
    }

    [Fact]
    public void Moved_formulas_keep_their_relative_references_and_formulas_outside_are_unchanged()
    {
        var s = Sheet(out var w);
        double[] units = [3, 1, 2];
        for (var i = 0; i < 3; i++)
        {
            s[i + 2, 1].Value = units[i];
            s[i + 2, 2].Formula = $"=A{i + 2}*10+$E$1";
        }

        s["E1"].Value = 0.5;
        s["D1"].Formula = "=B2";
        w.Recalculate();

        s.Range("A2:B4").Sort();
        w.Recalculate();

        Assert.Equal("1,2,3", Column(s, 1, 2, 4));
        Assert.Equal(["=A2*10+$E$1", "=A3*10+$E$1", "=A4*10+$E$1"], Enumerable.Range(2, 3).Select(row => s[row, 2].Formula));
        Assert.Equal("10.5,20.5,30.5", Column(s, 2, 2, 4));
        Assert.Equal(("=B2", 10.5), (s["D1"].Formula, s["D1"].Value.AsNumber()));
    }

    [Fact]
    public void Formulas_out_of_date_are_calculated_before_sorting_by_them()
    {
        var s = Sheet(out var w);
        s["A1"].Value = 1;
        s["A2"].Value = 2;
        s["B1"].Formula = "=A1*-1";
        s["B2"].Formula = "=A2*-1";

        s.Range("A1:B2").Sort(new SortKey(2));
        w.Recalculate();

        Assert.Equal("2,1", Column(s, 1, 1, 2));
    }

    [Fact]
    public void Formats_styles_links_and_notes_move_with_their_cells_but_validation_stays()
    {
        var s = Sheet(out var w);
        s["A1"].Value = 2;
        s["A2"].Value = 1;
        s["A1"].NumberFormat = "0.00";
        s["A1"].Style = CellStyle.Default with { Font = CellFont.Default with { Bold = true } };
        s["A1"].Hyperlink = new Hyperlink("https://example.com");
        s["A1"].Note = new Note("two");
        s.AddValidation("A1", DataValidation.ListFrom("$H$1:$H$2"));
        w.Recalculate();

        s.Range("A1:A2").Sort();

        Assert.Equal(("1", "2.00"), (s["A1"].Text, s["A2"].Text));
        Assert.True(s["A2"].Style.Font.Bold);
        Assert.False(s["A1"].Style.Font.Bold);
        Assert.Equal("https://example.com", s["A2"].Hyperlink?.Target);
        Assert.Null(s["A1"].Hyperlink);
        Assert.Equal("two", s["A2"].Note?.Text);
        Assert.Equal("A1", Assert.Single(s.Validations).Range.Address);
    }

    [Fact]
    public void Left_to_right_sorts_columns_by_a_row()
    {
        var s = Sheet(out var w);
        s["A1"].Value = 3;
        s["B1"].Value = 1;
        s["C1"].Value = 2;
        s["A2"].Value = "c";
        s["B2"].Value = "a";
        s["C2"].Value = "b";
        w.Recalculate();

        s.Range("A1:C2").Sort([new SortKey(1)], leftToRight: true);

        Assert.Equal(["a", "b", "c"], new[] { "A2", "B2", "C2" }.Select(a => s[a].Text));
    }

    [Fact]
    public void Whole_columns_sort_the_rows_used()
    {
        var s = Sheet(out var w);
        s["A1"].Value = 2;
        s["A2"].Value = 1;
        w.Recalculate();

        s.Range("A:B").Sort();

        Assert.Equal("1,2", Column(s, 1, 1, 2));
    }

    [Fact]
    public void Merged_cells_arrays_and_keys_outside_the_range_are_refused()
    {
        var s = Sheet(out var w);
        s["A1"].Value = 1;
        s["A2"].Value = 2;
        s.Merge("B1:B2");
        s["D1"].Formula = "=SEQUENCE(2)";
        w.Recalculate();

        Assert.Throws<InvalidOperationException>(() => s.Range("A1:B2").Sort());
        Assert.Throws<InvalidOperationException>(() => s.Range("D1:D2").Sort());
        Assert.Throws<ArgumentOutOfRangeException>(() => s.Range("A1:A2").Sort(new SortKey(2)));
    }

    [Fact]
    public void A_sorted_range_is_saved_and_reads_back()
    {
        var workbook = XlsxWriterTests.Load(new TestXlsx().Sheet("S", """
            <row r="1"><c r="A1"><v>3</v></c><c r="B1"><f>A1*2</f><v>6</v></c></row>
            <row r="2"><c r="A2"><v>1</v></c><c r="B2"><f>A2*2</f><v>2</v></c></row>
            """));
        var sheet = workbook["S"];

        sheet.Range("A1:B2").Sort();
        var reloaded = XlsxWriterTests.Reload(XlsxWriterTests.Save(Recalculated(workbook)))["S"];

        Assert.Equal(["=A1*2", "=A2*2"], new[] { "B1", "B2" }.Select(a => reloaded[a].Formula));
        Assert.Equal([1.0, 2, 3, 6], new[] { "A1", "B1", "A2", "B2" }.Select(a => reloaded[a].Value.AsNumber()));
    }

    private static Workbook Recalculated(Workbook workbook)
    {
        workbook.Recalculate();
        return workbook;
    }
}
