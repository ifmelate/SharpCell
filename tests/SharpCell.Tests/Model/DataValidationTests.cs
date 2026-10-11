using SharpCell.Parsing;

namespace SharpCell.Tests.Model;

/// <summary>Data validation rules: what a cell accepts, as Excel's Data Validation dialog sets it.</summary>
public class DataValidationTests
{
    private static Worksheet NewSheet() => new Workbook().AddSheet("S");

    private static DataValidation WholeNumberBetween(int low, int high) => new()
    {
        Type = ValidationType.WholeNumber,
        Operator = ValidationOperator.Between,
        Formula1 = "=" + low,
        Formula2 = "=" + high,
    };

    private static List<(string Range, DataValidation Validation)> Listed(Worksheet sheet) =>
        sheet.Validations.Select(v => (v.Range.Address, v.Validation)).ToList();

    [Fact]
    public void A_new_validation_allows_anything_and_shows_its_messages()
    {
        var validation = new DataValidation();

        Assert.Equal(ValidationType.Any, validation.Type);
        Assert.Equal(ValidationOperator.Between, validation.Operator);
        Assert.Null(validation.Formula1);
        Assert.Null(validation.Formula2);
        Assert.True(validation.AllowBlank);
        Assert.True(validation.ShowDropDown);
        Assert.True(validation.ShowInputMessage);
        Assert.True(validation.ShowErrorMessage);
        Assert.Equal(ValidationErrorStyle.Stop, validation.ErrorStyle);
        Assert.Null(validation.InputTitle);
        Assert.Null(validation.ErrorMessage);
    }

    [Fact]
    public void A_list_of_items_is_one_quoted_literal()
    {
        var validation = DataValidation.List("Yes", "No");

        Assert.Equal(ValidationType.List, validation.Type);
        Assert.Equal("=\"Yes,No\"", validation.Formula1);
        Assert.Equal("=\"Say \"\"hi\"\",Bye\"", DataValidation.List("Say \"hi\"", "Bye").Formula1);
    }

    [Fact]
    public void A_list_item_with_a_comma_or_a_list_over_255_characters_is_refused()
    {
        Assert.Throws<ArgumentException>(() => DataValidation.List("Yes", "No, thanks"));
        Assert.Throws<ArgumentException>(() => DataValidation.List(new string('x', 200), new string('y', 55)));
        Assert.Equal(258, DataValidation.List(new string('x', 200), new string('y', 54)).Formula1!.Length);
        Assert.Throws<ArgumentException>(() => DataValidation.List());
    }

    [Fact]
    public void A_list_from_a_range_and_a_custom_formula_keep_their_text()
    {
        Assert.Equal(new DataValidation { Type = ValidationType.List, Formula1 = "=$H$1:$H$5" }, DataValidation.ListFrom("$H$1:$H$5"));
        Assert.Equal(new DataValidation { Type = ValidationType.Custom, Formula1 = "=A1>0" }, DataValidation.Custom("=A1>0"));
        Assert.Equal("=A1>0", DataValidation.Custom("A1>0").Formula1);
    }

    [Fact]
    public void A_copy_takes_the_rule_and_the_changes_of_its_initializer()
    {
        var rule = new DataValidation(DataValidation.List("Yes", "No")) { ErrorTitle = "Pick one", ErrorMessage = "Choose Yes or No." };

        Assert.Equal(ValidationType.List, rule.Type);
        Assert.Equal("=\"Yes,No\"", rule.Formula1);
        Assert.Equal("Pick one", rule.ErrorTitle);
        Assert.Equal("Choose Yes or No.", rule.ErrorMessage);
    }

    [Fact]
    public void A_copy_without_changes_equals_its_source()
    {
        var source = new DataValidation
        {
            Type = ValidationType.Date, Operator = ValidationOperator.NotBetween, Formula1 = "=1", Formula2 = "=2",
            AllowBlank = false, ShowDropDown = false, ShowInputMessage = false, InputTitle = "t", InputMessage = "m",
            ShowErrorMessage = false, ErrorStyle = ValidationErrorStyle.Information, ErrorTitle = "e", ErrorMessage = "x",
        };

        Assert.Equal(source, new DataValidation(source));
        Assert.Throws<ArgumentNullException>(() => new DataValidation(null!));
        Assert.Throws<ArgumentException>(() => new DataValidation(source) { InputTitle = new string('x', 33) });
    }

    [Fact]
    public void A_title_over_32_or_a_message_over_255_characters_is_refused()
    {
        Assert.Throws<ArgumentException>(() => new DataValidation { InputTitle = new string('x', 33) });
        Assert.Throws<ArgumentException>(() => new DataValidation { ErrorTitle = new string('x', 33) });
        Assert.Throws<ArgumentException>(() => new DataValidation { InputMessage = new string('x', 256) });
        Assert.Throws<ArgumentException>(() => new DataValidation { ErrorMessage = new string('x', 256) });

        var longest = new DataValidation
        {
            InputTitle = new string('x', 32), ErrorTitle = new string('x', 32),
            InputMessage = new string('x', 255), ErrorMessage = new string('x', 255),
        };
        Assert.Equal(32, longest.InputTitle!.Length);
    }

    [Fact]
    public void Validations_with_the_same_settings_are_equal()
    {
        Assert.Equal(WholeNumberBetween(1, 10), WholeNumberBetween(1, 10));
        Assert.Equal(WholeNumberBetween(1, 10).GetHashCode(), WholeNumberBetween(1, 10).GetHashCode());
        Assert.NotEqual(WholeNumberBetween(1, 10), WholeNumberBetween(1, 11));
        Assert.NotEqual(new DataValidation { ErrorTitle = "a" }, new DataValidation { ErrorTitle = "A" });
        Assert.NotEqual(new DataValidation(), new DataValidation { ShowDropDown = false });
        Assert.False(new DataValidation().Equals(null));
    }

    [Fact]
    public void A_validation_applies_to_the_cells_of_its_range()
    {
        var sheet = NewSheet();
        var rule = WholeNumberBetween(1, 10);

        sheet.AddValidation("B2:B10", rule);

        Assert.Equal(rule, sheet["B5"].Validation);
        Assert.Equal(rule, sheet["B2"].Validation);
        Assert.Null(sheet["B11"].Validation);
        Assert.Null(sheet["C5"].Validation);
        Assert.Equal([("B2:B10", rule)], Listed(sheet));
    }

    [Fact]
    public void A_second_validation_replaces_the_first_on_the_cells_they_share()
    {
        var sheet = NewSheet();
        var first = WholeNumberBetween(1, 10);
        var second = DataValidation.List("Yes", "No");
        sheet.AddValidation("B2:B10", first);

        sheet.AddValidation("B5:C5", second);

        Assert.Equal([("B2:B4", first), ("B6:B10", first), ("B5:C5", second)], Listed(sheet));
        Assert.Equal(second, sheet["B5"].Validation);
        Assert.Equal(first, sheet["B6"].Validation);
    }

    [Fact]
    public void Clearing_takes_cells_out_of_every_rule()
    {
        var sheet = NewSheet();
        sheet.AddValidation("B2:B10", WholeNumberBetween(1, 10));
        sheet.AddValidation("D2:F4", DataValidation.List("Yes", "No"));

        sheet.ClearValidation("E3");
        Assert.Equal(["B2:B10", "D2:F2", "D3", "F3", "D4:F4"], Listed(sheet).Select(v => v.Range));
        Assert.Null(sheet["E3"].Validation);

        sheet.ClearValidation("A1:Z100");
        Assert.Empty(sheet.Validations);
    }

    [Fact]
    public void A_split_rule_keeps_its_formulas_relative_to_its_new_first_cell()
    {
        var sheet = NewSheet();
        sheet.AddValidation("A5:A10", DataValidation.Custom("=AND(A5>0,$C$1<A5)"));

        sheet.ClearValidation("A5:A6");

        var (range, validation) = Assert.Single(sheet.Validations);
        Assert.Equal("A7:A10", range.Address);
        Assert.Equal("=AND(A7>0,$C$1<A7)", validation.Formula1);
        Assert.Equal(validation, sheet["A9"].Validation);
    }

    [Fact]
    public void Formulas_follow_the_first_cell_of_the_rules_first_range()
    {
        var sheet = NewSheet();
        sheet.AddValidation("B2:C3", new DataValidation
        {
            Type = ValidationType.Decimal, Operator = ValidationOperator.LessThan, Formula1 = "=A2",
        });

        sheet.ClearValidation("B2");

        Assert.Equal(["C2", "B3:C3"], Listed(sheet).Select(v => v.Range));
        Assert.Equal("=B2", sheet.Validations[0].Validation.Formula1);
    }

    [Fact]
    public void A_type_that_compares_needs_its_formulas()
    {
        var sheet = NewSheet();

        Assert.Throws<ArgumentException>(() => sheet.AddValidation("A1", new DataValidation { Type = ValidationType.WholeNumber, Formula1 = "=1" }));
        Assert.Throws<ArgumentException>(() => sheet.AddValidation("A1", new DataValidation
        {
            Type = ValidationType.Date, Operator = ValidationOperator.NotBetween, Formula1 = "=1",
        }));
        Assert.Throws<ArgumentException>(() => sheet.AddValidation("A1", new DataValidation { Type = ValidationType.TextLength, Operator = ValidationOperator.LessThan }));
        Assert.Throws<ArgumentException>(() => sheet.AddValidation("A1", new DataValidation { Type = ValidationType.List }));
        Assert.Empty(sheet.Validations);

        sheet.AddValidation("A1", new DataValidation { Type = ValidationType.TextLength, Operator = ValidationOperator.LessThan, Formula1 = "=5" });
        sheet.AddValidation("A2", DataValidation.List("a"));
        sheet.AddValidation("A3", new DataValidation { InputMessage = "Anything goes" });
        Assert.Equal(3, sheet.Validations.Count);
    }

    [Fact]
    public void A_formula_that_cannot_be_parsed_adds_nothing()
    {
        var sheet = NewSheet();
        sheet.AddValidation("B2:B10", WholeNumberBetween(1, 10));

        Assert.Throws<FormulaParseException>(() => sheet.AddValidation("B5", DataValidation.Custom("=SUM(")));

        Assert.Equal(["B2:B10"], Listed(sheet).Select(v => v.Range));
    }

    [Fact]
    public void A_bad_range_or_a_null_validation_is_refused()
    {
        var sheet = NewSheet();

        Assert.Throws<ArgumentException>(() => sheet.AddValidation("nope", new DataValidation()));
        Assert.Throws<ArgumentNullException>(() => sheet.AddValidation("A1", null!));
        Assert.Throws<ArgumentException>(() => sheet.ClearValidation("nope"));
    }

    [Fact]
    public void A_clone_keeps_the_validations_and_is_independent()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        var rule = DataValidation.Custom("=A1>0");
        sheet.AddValidation("A1:A10", rule);

        var clone = workbook.Clone();
        Assert.Equal([("A1:A10", rule)], Listed(clone["S"]));

        clone["S"].ClearValidation("A1:A2");
        clone["S"].AddValidation("C1", DataValidation.List("x"));
        Assert.Equal([("A1:A10", rule)], Listed(sheet));
        Assert.Equal("=A3>0", clone["S"].Validations[0].Validation.Formula1);
    }

    [Fact]
    public void Changing_validations_from_a_custom_function_is_refused()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        var caught = new List<Exception>();
        workbook.Functions.Add("HOOK", _ =>
        {
            try { sheet.AddValidation("B1", DataValidation.List("a")); } catch (Exception ex) { caught.Add(ex); }
            try { sheet.ClearValidation("C1"); } catch (Exception ex) { caught.Add(ex); }
            return 5;
        });
        sheet.AddValidation("C1", DataValidation.List("b"));
        sheet["A1"].Formula = "=HOOK()";
        workbook.Recalculate();

        Assert.All(caught, ex => Assert.IsType<InvalidOperationException>(ex));
        Assert.Equal(2, caught.Count);
        Assert.Equal(["C1"], Listed(sheet).Select(v => v.Range));
    }

    [Fact]
    public void A_rule_read_from_a_file_keeps_a_formula_SharpCell_cannot_read()
    {
        var sheet = NewSheet();
        var rule = DataValidation.Custom("=SUM(");

        sheet.AddValidationRule([new(1, 1, 5, 1), new(1, 3, 5, 3)], rule);

        Assert.Equal([("A1:A5", rule), ("C1:C5", rule)], Listed(sheet));
        Assert.IsType<UnsupportedNode>(sheet.ValidationRules[0].F1);
    }

    [Fact]
    public void The_first_cell_of_a_rule_SharpCell_cannot_read_stays_put()
    {
        var sheet = NewSheet();
        var rule = DataValidation.Custom("=SUM(");
        sheet.AddValidationRule([new(1, 1, 5, 1)], rule);

        Assert.Throws<NotSupportedException>(() => sheet.ClearValidation("A1"));
        Assert.Throws<NotSupportedException>(() => sheet.AddValidation("A1:B1", DataValidation.List("x")));
        Assert.Equal([("A1:A5", rule)], Listed(sheet));

        sheet.ClearValidation("A3");
        Assert.Equal([("A1:A2", rule), ("A4:A5", rule)], Listed(sheet));
        sheet.ClearValidation("A1:A5");
        Assert.Empty(sheet.Validations);
    }

    [Fact]
    public void Validations_follow_deleted_rows()
    {
        var sheet = NewSheet();
        sheet.AddValidation("A5:A10", DataValidation.Custom("=A5>0"));

        sheet.DeleteRows(1, 2);

        var (range, validation) = Assert.Single(sheet.Validations);
        Assert.Equal("A3:A8", range.Address);
        Assert.Equal("=A3>0", validation.Formula1);
        Assert.Equal(validation, sheet["A8"].Validation);
        Assert.Null(sheet["A9"].Validation);

        sheet.DeleteRows(3, 6);
        Assert.Empty(sheet.Validations);
    }

    [Fact]
    public void Clearing_the_first_cells_moves_the_formulas_with_the_rule()
    {
        var sheet = NewSheet();
        sheet.AddValidation("A5:A10", DataValidation.Custom("=A5>0"));
        sheet.DeleteRows(1, 2);

        sheet.ClearValidation("A3:A4");

        var (range, validation) = Assert.Single(sheet.Validations);
        Assert.Equal("A5:A8", range.Address);
        Assert.Equal("=A5>0", validation.Formula1);
    }

    [Fact]
    public void Inserted_columns_move_the_rule_and_the_list_it_takes_items_from()
    {
        var sheet = NewSheet();
        sheet.AddValidation("B1:B10", DataValidation.ListFrom("=$H$1:$H$5"));

        sheet.InsertColumns(1);

        var (range, validation) = Assert.Single(sheet.Validations);
        Assert.Equal("C1:C10", range.Address);
        Assert.Equal("=$I$1:$I$5", validation.Formula1);
        Assert.Equal(ValidationType.List, validation.Type);
        Assert.Null(sheet["B1"].Validation);
    }

    [Fact]
    public void A_rule_naming_another_sheets_cells_follows_that_sheets_edits()
    {
        var workbook = new Workbook();
        var form = workbook.AddSheet("Form");
        var lists = workbook.AddSheet("Lists");
        form.AddValidation("A1:A5", DataValidation.ListFrom("=Lists!$A$2:$A$4"));
        form.AddValidation("B2:B5", DataValidation.Custom("=B2<Lists!C2"));

        lists.InsertRows(1, 2);
        lists.DeleteColumns(1);

        Assert.Equal(["A1:A5", "B2:B5"], Listed(form).Select(v => v.Range));
        Assert.Equal("=Lists!#REF!", form.Validations[0].Validation.Formula1);
        Assert.Equal("=B2<Lists!B4", form.Validations[1].Validation.Formula1);
    }

    [Fact]
    public void A_rule_whose_first_cells_are_deleted_still_checks_each_cell_as_before()
    {
        var sheet = NewSheet();
        sheet.AddValidation("A1:A10", DataValidation.Custom("=AND(A1>0,$C$9>A1)"));
        sheet.AddValidationRule([new(1, 4, 2, 4), new(5, 6, 6, 6)], DataValidation.Custom("=D1<>\"\""));

        sheet.DeleteRows(1, 2);

        Assert.Equal(["A1:A8", "F3:F4"], Listed(sheet).Select(v => v.Range));
        Assert.Equal("=AND(A1>0,$C$7>A1)", sheet.Validations[0].Validation.Formula1);
        Assert.Equal("=F3<>\"\"", sheet.Validations[1].Validation.Formula1);
    }

    [Fact]
    public void A_rule_whose_first_range_is_deleted_takes_its_formulas_to_the_next_range()
    {
        var columns = NewSheet();
        columns.AddValidationRule([new(1, 3, 2, 3), new(5, 1, 6, 1)], DataValidation.Custom("=C1>0"));
        var rows = NewSheet();
        rows.AddValidationRule([new(10, 1, 11, 1), new(1, 6, 2, 6)], DataValidation.Custom("=A10>0"));

        columns.DeleteColumns(3);
        rows.DeleteRows(10, 2);

        Assert.Equal([("A5:A6", "=A5>0")], Listed(columns).Select(v => (v.Range, v.Validation.Formula1)));
        Assert.Equal([("F1:F2", "=F1>0")], Listed(rows).Select(v => (v.Range, v.Validation.Formula1)));
    }

    [Fact]
    public void Deleting_a_rules_first_column_moves_its_formulas_to_the_next_column()
    {
        var sheet = NewSheet();
        sheet.AddValidation("B2:D5", DataValidation.Custom("=AND(B2>0,$F$1>B2)"));

        sheet.DeleteColumns(2);

        var (range, validation) = Assert.Single(sheet.Validations);
        Assert.Equal("B2:C5", range.Address);
        Assert.Equal("=AND(B2>0,$E$1>B2)", validation.Formula1);
    }

    [Fact]
    public void Rows_inserted_inside_a_rule_grow_it_and_rows_above_move_it()
    {
        var sheet = NewSheet();
        sheet.AddValidation("A2:A5", WholeNumberBetween(1, 10));
        sheet.AddValidation("C2:C3", DataValidation.Custom("=C2>B2"));

        sheet.InsertRows(3, 2);
        Assert.Equal(["A2:A7", "C2:C5"], Listed(sheet).Select(v => v.Range));
        Assert.Equal(WholeNumberBetween(1, 10), sheet.Validations[0].Validation);

        sheet.InsertRows(1);
        Assert.Equal(["A3:A8", "C3:C6"], Listed(sheet).Select(v => v.Range));
        Assert.Equal("=C3>B3", sheet.Validations[1].Validation.Formula1);
    }

    [Fact]
    public void A_rule_SharpCell_cannot_read_on_the_sheet_or_naming_it_blocks_edits()
    {
        var workbook = new Workbook();
        var data = workbook.AddSheet("Data");
        var other = workbook.AddSheet("Other");
        data["A1"].Value = 1;
        data.AddValidationRule([new(1, 2, 1, 2)], DataValidation.Custom("=SUM("));

        Assert.Throws<NotSupportedException>(() => data.InsertRows(5));
        Assert.Equal(["B1"], Listed(data).Select(v => v.Range));
        other.InsertRows(1);

        data.ClearValidation("B1");
        other.AddValidationRule([new(1, 1, 1, 1)], DataValidation.Custom("=SUM(Data!A1"));
        Assert.Throws<NotSupportedException>(() => data.DeleteColumns(3));
        Assert.Equal(1, data["A1"].Value);

        other.AddValidationRule([new(2, 1, 2, 1)], DataValidation.Custom("=SUM("));
        other.ClearValidation("A1");
        data.DeleteColumns(3);
    }

    [Fact]
    public void A_rule_pushed_past_the_last_row_is_refused_but_a_whole_column_rule_is_cut()
    {
        var sheet = NewSheet();
        sheet.AddValidation("B1048576", DataValidation.List("x"));

        Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));
        Assert.Equal(["B1048576"], Listed(sheet).Select(v => v.Range));

        sheet.ClearValidation("B1048576");
        sheet.AddValidation("A2:A1048576", DataValidation.List("y"));
        sheet.InsertRows(1);
        Assert.Equal(["A3:A1048576"], Listed(sheet).Select(v => v.Range));
    }
}
