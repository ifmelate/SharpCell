using System;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Tests.Model;

public class CopyCellsTests
{
    private static (Workbook Workbook, Worksheet Sheet) NewSheet()
    {
        var workbook = new Workbook();
        return (workbook, workbook.AddSheet("S"));
    }

    private static CellValue N(double value) => CellValue.Number(value);

    [Fact]
    public void Relative_references_keep_their_distance_and_absolute_ones_stay()
    {
        var (workbook, sheet) = NewSheet();
        sheet["A1"].Value = N(1);
        sheet["A2"].Value = N(2);
        sheet["B1"].Formula = "=A1*$A$1+A$1+$A1";

        var written = sheet.Range("B1").CopyTo(sheet.Range("B2"));
        workbook.Recalculate();

        Assert.Equal("B2", written.Address);
        Assert.Equal("=A2*$A$1+A$1+$A2", sheet["B2"].Formula);
        Assert.Equal(N(5), sheet["B2"].Value);
    }

    [Fact]
    public void A_reference_past_the_sheet_edge_becomes_REF()
    {
        var (workbook, sheet) = NewSheet();
        sheet["B2"].Formula = "=A1+Other!A1+SUM(A1:B1)+$A$1";
        workbook.AddSheet("Other");

        sheet.Range("B2").CopyTo(sheet.Range("A1"));

        Assert.Equal("=#REF!+Other!#REF!+SUM(#REF!)+$A$1", sheet["A1"].Formula);
    }

    [Fact]
    public void Whole_rows_and_columns_move_on_their_own_axis()
    {
        var (_, sheet) = NewSheet();
        sheet["C3"].Formula = "=SUM(1:1)+SUM(A:A)";

        sheet.Range("C3").CopyTo(sheet.Range("D4"));
        sheet.Range("C3").CopyTo(sheet.Range("C1"));

        Assert.Equal("=SUM(2:2)+SUM(B:B)", sheet["D4"].Formula);
        Assert.Equal("=SUM(#REF!)+SUM(A:A)", sheet["C1"].Formula);
    }

    [Fact]
    public void Values_number_formats_and_styles_are_copied()
    {
        var (_, sheet) = NewSheet();
        var bold = CellStyle.Default with { Font = CellFont.Default with { Bold = true } };
        sheet["A1"].Value = N(0.5);
        sheet["A1"].NumberFormat = "0%";
        sheet["A1"].Style = bold;
        sheet["A2"].Style = bold;

        sheet.Range("A1:A2").CopyTo(sheet.Range("C5"));

        Assert.Equal(N(0.5), sheet["C5"].Value);
        Assert.Equal("50%", sheet["C5"].Text);
        Assert.Equal(bold, sheet["C5"].Style);
        Assert.Equal(bold, sheet["C6"].Style);
        Assert.Equal(CellValue.Empty, sheet["C6"].Value);
    }

    [Fact]
    public void Empty_cells_of_the_range_clear_what_they_are_copied_over()
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Value = N(1);
        sheet["C1"].Value = N(9);
        sheet["C2"].Formula = "=C1";
        sheet["C2"].NumberFormat = "0.00";
        sheet["C2"].Style = CellStyle.Default with { Fill = CellColor.FromRgb(0xFF0000) };

        sheet.Range("A1:A2").CopyTo(sheet.Range("C1"));

        Assert.Equal(N(1), sheet["C1"].Value);
        Assert.Null(sheet["C2"].Formula);
        Assert.Equal(CellValue.Empty, sheet["C2"].Value);
        Assert.Equal("General", sheet["C2"].NumberFormat);
        Assert.Equal(sheet.Workbook.DefaultStyle, sheet["C2"].Style);
    }

    [Fact]
    public void A_destination_a_whole_number_of_times_the_range_is_filled_by_repeating_it()
    {
        var (workbook, sheet) = NewSheet();
        sheet["A1"].Value = N(1);
        sheet["A2"].Value = N(2);
        sheet["B1"].Formula = "=A1*10";
        sheet["B2"].Formula = "=A2*10";

        var written = sheet.Range("B1:B2").CopyTo(sheet.Range("C1:D4"));
        workbook.Recalculate();

        Assert.Equal("C1:D4", written.Address);
        Assert.Equal("=B1*10", sheet["C1"].Formula);
        Assert.Equal("=C4*10", sheet["D4"].Formula);
        Assert.Equal(N(100), sheet["C1"].Value);
        Assert.Equal(N(200), sheet["C2"].Value);
    }

    [Fact]
    public void A_destination_that_is_no_multiple_takes_one_copy_at_its_top_left_cell()
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Value = N(1);
        sheet["A2"].Value = N(2);

        var written = sheet.Range("A1:A2").CopyTo(sheet.Range("C1:C3"));

        Assert.Equal("C1:C2", written.Address);
        Assert.Equal(CellValue.Empty, sheet["C3"].Value);
    }

    [Fact]
    public void Overlapping_ranges_copy_the_cells_as_they_were()
    {
        var (workbook, sheet) = NewSheet();
        sheet["A1"].Value = N(1);
        sheet["A2"].Formula = "=A1+1";
        sheet["A3"].Formula = "=A2+1";

        sheet.Range("A1:A3").CopyTo(sheet.Range("A2"));
        workbook.Recalculate();

        Assert.Equal(N(1), sheet["A2"].Value);
        Assert.Equal("=A2+1", sheet["A3"].Formula);
        Assert.Equal("=A3+1", sheet["A4"].Formula);
        Assert.Equal(N(3), sheet["A4"].Value);
    }

    [Fact]
    public void A_spilling_formula_copied_with_its_cells_spills_again_and_spilled_cells_alone_copy_as_values()
    {
        var (workbook, sheet) = NewSheet();
        sheet["A1"].Formula = "=SEQUENCE(3)";
        workbook.Recalculate();

        sheet.Range("A1:A3").CopyTo(sheet.Range("C1"));
        sheet.Range("A2:A3").CopyTo(sheet.Range("E1"));
        workbook.Recalculate();

        Assert.Equal("=SEQUENCE(3)", sheet["C1"].Formula);
        Assert.Null(sheet["C2"].Formula);
        Assert.Equal(N(3), sheet["C3"].Value);
        Assert.Null(sheet["E1"].Formula);
        Assert.Equal([N(2), N(3)], [sheet["E1"].Value, sheet["E2"].Value]);
    }

    [Fact]
    public void Copying_to_another_workbook_reads_formulas_there()
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Formula = "=Data!A1*2";
        var other = new Workbook();
        var data = other.AddSheet("Data");
        var target = other.AddSheet("T");
        data["B2"].Value = N(21);

        sheet.Range("A1").CopyTo(target.Range("B2"));
        other.Recalculate();

        Assert.Equal("=Data!B2*2", target["B2"].Formula);
        Assert.Equal(N(42), target["B2"].Value);
    }

    [Fact]
    public void Formulas_reading_the_written_cells_calculate_again()
    {
        var (workbook, sheet) = NewSheet();
        sheet["A1"].Value = N(5);
        sheet["D1"].Formula = "=C1+1";
        workbook.Recalculate();

        sheet.Range("A1").CopyTo(sheet.Range("C1"));
        var result = workbook.Recalculate();

        Assert.Equal(N(6), sheet["D1"].Value);
        Assert.Contains(result.ChangedCells, c => c.Address == "D1");
    }

    [Fact]
    public void A_copy_past_the_sheet_edge_is_refused()
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Value = N(1);

        Assert.Throws<ArgumentException>(() => sheet.Range("A1:A2").CopyTo(sheet.Range("A1048576")));
        Assert.Equal(CellValue.Empty, sheet["A1048576"].Value);
    }

    [Fact]
    public void Array_formulas_are_refused_before_anything_changes()
    {
        var (_, sheet) = NewSheet();
        var loader = new SheetLoader(sheet);
        loader.SetFormula(1, 1, FormulaParser.Parse("={1;2}", new CellAddress(1, 1)), LoadedFormulaKind.Array, new Area(1, 1, 2, 1), N(1));
        loader.Complete();
        sheet.Workbook.Recalculate();
        sheet["C1"].Value = N(7);

        Assert.Throws<NotSupportedException>(() => sheet.Range("A1:A2").CopyTo(sheet.Range("C1")));
        Assert.Throws<InvalidOperationException>(() => sheet.Range("C1:C2").CopyTo(sheet.Range("A1")));
        Assert.Equal(N(7), sheet["C1"].Value);
        Assert.Equal(N(2), sheet["A2"].Value);
    }

    [Fact]
    public void Hyperlinks_are_copied_and_cells_without_one_clear_theirs()
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Hyperlink = new Hyperlink("https://example.com");
        sheet["C2"].Hyperlink = new Hyperlink("#S!A1");

        sheet.Range("A1:A2").CopyTo(sheet.Range("C1"));

        Assert.Equal("https://example.com", sheet["C1"].Hyperlink?.Target);
        Assert.Null(sheet["C2"].Hyperlink);
    }

    [Fact]
    public void Data_validation_is_copied_with_its_formulas_following_the_cells()
    {
        var (_, sheet) = NewSheet();
        sheet.AddValidation("A1:A2", DataValidation.List("Yes", "No"));
        sheet.AddValidation("B1", DataValidation.Custom("=B1>A1"));
        sheet.AddValidation("D1:D5", DataValidation.Custom("=D1<>0"));

        sheet.Range("A1:B1").CopyTo(sheet.Range("C3:D4"));

        Assert.Equal(DataValidation.List("Yes", "No"), sheet["C4"].Validation);
        Assert.Equal("=D3>C3", sheet["D3"].Validation?.Formula1);
        Assert.Equal("=D4>C4", sheet["D4"].Validation?.Formula1);
        Assert.Equal("=D1<>0", sheet["D5"].Validation?.Formula1);
        Assert.Equal(DataValidation.List("Yes", "No"), sheet["A2"].Validation);
    }

    [Fact]
    public void Cells_without_validation_clear_it_where_they_are_copied()
    {
        var (_, sheet) = NewSheet();
        sheet.AddValidation("B1:B3", DataValidation.List("Yes", "No"));

        sheet.Range("A1").CopyTo(sheet.Range("B2"));

        Assert.Null(sheet["B2"].Validation);
        Assert.NotNull(sheet["B1"].Validation);
        Assert.NotNull(sheet["B3"].Validation);
    }

    [Fact]
    public void A_rule_over_several_ranges_is_added_as_Excel_writes_them()
    {
        var (_, sheet) = NewSheet();

        sheet.AddValidation("B2:B3 D2:D3", DataValidation.Custom("=B2>0"));

        Assert.Equal(["B2:B3", "D2:D3"], sheet.Validations.Select(v => v.Range.Address));
        Assert.Same(sheet.Validations[0].Validation, sheet.Validations[1].Validation);

        // Left with D3 alone, the rule's formula is written for D3, its new first cell.
        sheet.ClearValidation("B2:B3 D2");
        var (range, rule) = Assert.Single(sheet.Validations);
        Assert.Equal("D3", range.Address);
        Assert.Equal("=D3>0", rule.Formula1);
    }
}
