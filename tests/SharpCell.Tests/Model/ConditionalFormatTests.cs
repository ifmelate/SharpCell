namespace SharpCell.Tests.Model;

/// <summary>Conditional formats: which cells each rule applies to, and how the cells then look.</summary>
public class ConditionalFormatTests
{
    private static readonly ConditionalStyle Red = ConditionalStyle.LightRedFill;
    private static readonly ConditionalStyle Bold = new() { Bold = true };

    // A1:A6 hold 1, 5, 5, 3, empty, 10.
    private static (Workbook W, Worksheet S) Book()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        double?[] values = [1, 5, 5, 3, null, 10];
        for (var i = 0; i < values.Length; i++)
        {
            if (values[i] is { } value)
                sheet[i + 1, 1].Value = value;
        }

        workbook.Recalculate();
        return (workbook, sheet);
    }

    private static string[] Matching(Worksheet sheet, string range) =>
        [.. sheet.Range(range).Cells.Select(c => c.Address).Concat(EmptyCells(sheet, range)).Where(a => sheet[a].ConditionalFormats.Count > 0).Order()];

    private static IEnumerable<string> EmptyCells(Worksheet sheet, string range)
    {
        var r = sheet.Range(range);
        for (var row = r.FirstRow; row <= r.LastRow; row++)
        {
            for (var column = r.FirstColumn; column <= r.LastColumn; column++)
            {
                if (sheet[row, column].Value.Kind == CellValueKind.Empty)
                    yield return sheet[row, column].Address;
            }
        }
    }

    [Fact]
    public void Cell_value_rules_compare_the_value_and_an_empty_cell_compares_as_zero()
    {
        var (_, sheet) = Book();
        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.CellValue(ConditionalOperator.LessThan, "=4", Red));

        Assert.Equal(["A1", "A4", "A5"], Matching(sheet, "A1:A6"));
    }

    [Fact]
    public void Between_takes_its_bounds_either_way_round()
    {
        var (_, sheet) = Book();
        sheet["C1"].Value = 6;
        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.CellValue(ConditionalOperator.Between, "=$C$1", Red, "=3"));

        Assert.Equal(["A2", "A3", "A4"], Matching(sheet, "A1:A6"));
    }

    [Fact]
    public void An_expression_is_read_from_each_cell_as_from_the_first()
    {
        var (w, sheet) = Book();
        for (var row = 1; row <= 6; row++)
            sheet[row, 2].Value = 4;
        w.Recalculate();
        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.Expression("=A1>$B1", Red));

        Assert.Equal(["A2", "A3", "A6"], Matching(sheet, "A1:A6"));
        Assert.Equal("=A1>$B1", sheet.ConditionalFormats.Single().Format.Formula1);
    }

    [Fact]
    public void Top_and_bottom_rank_the_numbers_of_the_ranges_with_ties()
    {
        var (_, sheet) = Book();

        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.Top(2, Red));
        Assert.Equal(["A2", "A3", "A6"], Matching(sheet, "A1:A6"));

        sheet.ClearConditionalFormats();
        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.Bottom(40, Red, percent: true));   // 40% of 5 numbers: 2
        Assert.Equal(["A1", "A4"], Matching(sheet, "A1:A6"));
    }

    [Fact]
    public void Average_rules_compare_with_the_mean_of_the_numbers()
    {
        var (_, sheet) = Book();   // mean 24 / 5 = 4.8

        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.AboveAverage(Red));
        Assert.Equal(["A2", "A3", "A6"], Matching(sheet, "A1:A6"));

        sheet.ClearConditionalFormats();
        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.AboveAverage(Red, standardDeviations: 1));   // sd 3.42
        Assert.Equal(["A6"], Matching(sheet, "A1:A6"));

        sheet.ClearConditionalFormats();
        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.BelowAverage(Red));
        Assert.Equal(["A1", "A4"], Matching(sheet, "A1:A6"));
    }

    [Fact]
    public void Duplicates_compare_text_ignoring_case_and_leave_empty_cells_out()
    {
        var (_, sheet) = Book();
        sheet["B1"].Value = "Pen";
        sheet["B2"].Value = "pen";
        sheet["B3"].Value = "Book";

        sheet.AddConditionalFormat("A1:B6", ConditionalFormat.DuplicateValues(Red));
        Assert.Equal(["A2", "A3", "B1", "B2"], Matching(sheet, "A1:B6"));

        sheet.ClearConditionalFormats();
        sheet.AddConditionalFormat("A1:B6", ConditionalFormat.UniqueValues(Red));
        Assert.Equal(["A1", "A4", "A6", "B3"], Matching(sheet, "A1:B6"));
    }

    [Fact]
    public void Text_blank_and_error_rules_use_the_formulas_Excel_writes()
    {
        var (w, sheet) = Book();
        sheet["B1"].Value = "Total sales";
        sheet["B2"].Value = "sales";
        sheet["B3"].Value = "  ";
        sheet["B4"].Formula = "=1/0";
        w.Recalculate();

        void Check(ConditionalFormat format, params string[] expected)
        {
            sheet.ClearConditionalFormats();
            sheet.AddConditionalFormat("B1:B5", format);
            Assert.Equal(expected, Matching(sheet, "B1:B5"));
        }

        Check(ConditionalFormat.ContainsText("SALES", Red), "B1", "B2");
        Check(ConditionalFormat.BeginsWith("total", Red), "B1");
        Check(ConditionalFormat.EndsWith("es", Red), "B1", "B2");
        Check(ConditionalFormat.NotContainsText("sales", Red), "B3", "B4", "B5");
        Check(ConditionalFormat.Blanks(Red), "B3", "B5");
        Check(ConditionalFormat.Errors(Red), "B4");
        Check(ConditionalFormat.NoErrors(Red), "B1", "B2", "B3", "B5");
    }

    [Fact]
    public void Date_rules_count_from_today()
    {
        var (w, sheet) = Book();
        sheet["C1"].Formula = "=TODAY()";
        sheet["C2"].Formula = "=TODAY()-1";
        sheet["C3"].Formula = "=TODAY()+40";
        w.Recalculate();

        sheet.AddConditionalFormat("C1:C3", ConditionalFormat.DateOccurring(TimePeriod.Yesterday, Red));
        sheet.AddConditionalFormat("C1:C3", ConditionalFormat.DateOccurring(TimePeriod.Last7Days, Bold));

        Assert.Equal([ConditionalFormatType.TimePeriod, ConditionalFormatType.TimePeriod], sheet["C2"].ConditionalFormats.Select(f => f.Type));
        Assert.Equal(TimePeriod.Last7Days, sheet["C1"].ConditionalFormats.Single().Period);
        Assert.Empty(sheet["C3"].ConditionalFormats);
    }

    [Fact]
    public void The_display_style_puts_higher_rules_over_lower_ones_and_stop_if_true_stops()
    {
        var (_, sheet) = Book();
        sheet["A6"].Style = CellStyle.Default with { Font = CellFont.Default with { Italic = true } };
        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.Top(1, new ConditionalStyle { Fill = CellColor.FromRgb(0x00FF00) }));
        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.CellValue(ConditionalOperator.GreaterThan, "=4", Red));

        var shown = sheet["A6"].DisplayStyle;
        Assert.Equal((CellColor.FromRgb(0x00FF00), CellColor.FromRgb(0x9C0006), true), (shown.Fill, shown.Font.Color, shown.Font.Italic));
        Assert.Equal(sheet["A1"].Style, sheet["A1"].DisplayStyle);

        sheet.ClearConditionalFormats();
        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.Top(1, Bold) with { StopIfTrue = true });
        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.CellValue(ConditionalOperator.GreaterThan, "=4", Red));
        Assert.Single(sheet["A6"].ConditionalFormats);
        Assert.Equal(2, sheet["A2"].ConditionalFormats.Count + 1);
    }

    [Fact]
    public void Rules_follow_inserted_rows_and_their_formulas_move()
    {
        var (w, sheet) = Book();
        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.Expression("=A1>$C$1", Red));

        sheet.InsertRows(1);
        w.Recalculate();

        var (ranges, format) = sheet.ConditionalFormats.Single();
        Assert.Equal(["A2:A7"], ranges.Select(r => r.Address));
        Assert.Equal("=A2>$C$2", format.Formula1);
        Assert.Equal(["A2", "A3", "A4", "A5", "A7"], Matching(sheet, "A1:A7"));   // > empty C2, as 0
    }

    [Fact]
    public void Clearing_part_of_a_rule_splits_its_ranges_and_copying_copies_rules()
    {
        var (_, sheet) = Book();
        sheet.AddConditionalFormat("A1:A6", ConditionalFormat.Expression("=A1>4", Red));

        sheet.ClearConditionalFormats("A3:A4");
        Assert.Equal(["A1:A2", "A5:A6"], sheet.ConditionalFormats.Single().Ranges.Select(r => r.Address));

        sheet.Range("A1:A2").CopyTo(sheet.Range("D1"));
        Assert.Equal(2, sheet.ConditionalFormats.Count);
        Assert.Equal("=D1>4", sheet.ConditionalFormats[1].Format.Formula1);
    }

    [Fact]
    public void Wrong_rules_are_refused()
    {
        Assert.Throws<ArgumentException>(() => ConditionalFormat.CellValue(ConditionalOperator.Between, "=1", Red));
        Assert.Throws<ArgumentException>(() => ConditionalFormat.CellValue(ConditionalOperator.Equal, "=1", Red, "=2"));
        Assert.Throws<ArgumentOutOfRangeException>(() => ConditionalFormat.Top(0, Red));
        Assert.Throws<ArgumentOutOfRangeException>(() => ConditionalFormat.Top(101, Red, percent: true));
        Assert.Throws<ArgumentException>(() => ConditionalFormat.ContainsText("", Red));
    }
}
