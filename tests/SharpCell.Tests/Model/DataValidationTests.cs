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
}
