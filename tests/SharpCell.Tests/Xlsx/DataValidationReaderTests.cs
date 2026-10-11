using System.Linq;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>Data validation rules are read from a sheet's &lt;dataValidations&gt; element.</summary>
public class DataValidationReaderTests
{
    private static Worksheet Load(string validations, string tail = "") =>
        XlsxReader.Load(new TestXlsx()
            .RawSheet("S", $"<sheetData/><dataValidations count=\"9\">{validations}</dataValidations>{tail}")
            .Build())["S"];

    private static List<(string Range, DataValidation Validation)> Listed(Worksheet sheet) =>
        [.. sheet.Validations.Select(v => (v.Range.Address, v.Validation))];

    [Fact]
    public void Every_attribute_is_read()
    {
        var sheet = Load("""
            <dataValidation type="decimal" operator="notBetween" allowBlank="1" showDropDown="1" showInputMessage="1" showErrorMessage="1"
              errorStyle="warning" errorTitle="Out of range" error="Between 1 and 2, please." promptTitle="Rate" prompt="A rate &amp; more" sqref="B2:B5">
              <formula1>1</formula1><formula2>$C$1</formula2></dataValidation>
            """);

        var (range, validation) = Assert.Single(Listed(sheet));
        Assert.Equal("B2:B5", range);
        Assert.Equal(new DataValidation
        {
            Type = ValidationType.Decimal,
            Operator = ValidationOperator.NotBetween,
            Formula1 = "=1",
            Formula2 = "=$C$1",
            AllowBlank = true,
            ShowDropDown = false,
            ShowInputMessage = true,
            InputTitle = "Rate",
            InputMessage = "A rate & more",
            ShowErrorMessage = true,
            ErrorStyle = ValidationErrorStyle.Warning,
            ErrorTitle = "Out of range",
            ErrorMessage = "Between 1 and 2, please.",
        }, validation);
        Assert.Same(validation, sheet["B4"].Validation);
    }

    [Fact]
    public void Absent_flags_take_the_file_defaults()
    {
        // The schema's defaults are off; showDropDown="1" hides the drop-down, so its absence shows it.
        var validation = Load("<dataValidation sqref=\"A1\"/>").Validations.Single().Validation;

        Assert.Equal(new DataValidation { AllowBlank = false, ShowInputMessage = false, ShowErrorMessage = false }, validation);
        Assert.True(validation.ShowDropDown);
        Assert.Equal(ValidationType.Any, validation.Type);
        Assert.Equal(ValidationOperator.Between, validation.Operator);
        Assert.Equal(ValidationErrorStyle.Stop, validation.ErrorStyle);
    }

    [Theory]
    [InlineData("none", ValidationType.Any)]
    [InlineData("whole", ValidationType.WholeNumber)]
    [InlineData("decimal", ValidationType.Decimal)]
    [InlineData("list", ValidationType.List)]
    [InlineData("date", ValidationType.Date)]
    [InlineData("time", ValidationType.Time)]
    [InlineData("textLength", ValidationType.TextLength)]
    [InlineData("custom", ValidationType.Custom)]
    public void Every_type_is_read(string text, ValidationType type) =>
        Assert.Equal(type, Load($"<dataValidation type=\"{text}\" sqref=\"A1\"><formula1>1</formula1></dataValidation>").Validations.Single().Validation.Type);

    [Theory]
    [InlineData("between", ValidationOperator.Between)]
    [InlineData("notBetween", ValidationOperator.NotBetween)]
    [InlineData("equal", ValidationOperator.Equal)]
    [InlineData("notEqual", ValidationOperator.NotEqual)]
    [InlineData("greaterThan", ValidationOperator.GreaterThan)]
    [InlineData("lessThan", ValidationOperator.LessThan)]
    [InlineData("greaterThanOrEqual", ValidationOperator.GreaterThanOrEqual)]
    [InlineData("lessThanOrEqual", ValidationOperator.LessThanOrEqual)]
    public void Every_operator_is_read(string text, ValidationOperator op) =>
        Assert.Equal(op, Load($"<dataValidation type=\"whole\" operator=\"{text}\" sqref=\"A1\"><formula1>1</formula1></dataValidation>").Validations.Single().Validation.Operator);

    [Theory]
    [InlineData("stop", ValidationErrorStyle.Stop)]
    [InlineData("warning", ValidationErrorStyle.Warning)]
    [InlineData("information", ValidationErrorStyle.Information)]
    public void Every_error_style_is_read(string text, ValidationErrorStyle style) =>
        Assert.Equal(style, Load($"<dataValidation errorStyle=\"{text}\" sqref=\"A1\"/>").Validations.Single().Validation.ErrorStyle);

    [Fact]
    public void A_rule_over_several_ranges_is_one_entry_per_range()
    {
        var sheet = Load("<dataValidation type=\"custom\" sqref=\"A1:A5 C1:C5\"><formula1>A1&gt;B1</formula1></dataValidation>");

        var listed = Listed(sheet);
        Assert.Equal(["A1:A5", "C1:C5"], listed.Select(v => v.Range));
        Assert.Same(listed[0].Validation, listed[1].Validation);
        Assert.Equal("=A1>B1", listed[0].Validation.Formula1);
        Assert.Same(listed[0].Validation, sheet["C3"].Validation);
        Assert.Null(sheet["B3"].Validation);
    }

    [Fact]
    public void Formulas_are_read_relative_to_the_first_cell_and_without_file_prefixes()
    {
        var sheet = Load("""
            <dataValidation type="list" sqref="D4"><formula1>"Yes,No"</formula1></dataValidation>
            <dataValidation type="custom" sqref="E5:E9"><formula1>_xlfn.ISOWEEKNUM(E5)&gt;0</formula1></dataValidation>
            <dataValidation type="list" sqref="F1"><formula1>Other!$A$1:$A$3</formula1></dataValidation>
            """);

        Assert.Equal(["=\"Yes,No\"", "=ISOWEEKNUM(E5)>0", "=Other!$A$1:$A$3"], sheet.Validations.Select(v => v.Validation.Formula1));
    }

    [Fact]
    public void A_damaged_rule_is_skipped_and_the_others_are_read()
    {
        var sheet = Load("""
            <dataValidation type="whole" sqref="A1:"><formula1>1</formula1></dataValidation>
            <dataValidation type="whole" sqref=""><formula1>1</formula1></dataValidation>
            <dataValidation type="whole"><formula1>1</formula1></dataValidation>
            <dataValidation type="money" sqref="A2"><formula1>1</formula1></dataValidation>
            <dataValidation type="whole" operator="around" sqref="A3"><formula1>1</formula1></dataValidation>
            <dataValidation errorStyle="loud" sqref="A4"/>
            <dataValidation type="list" sqref="B1"><formula1>"a,b"</formula1></dataValidation>
            """);

        Assert.Equal(["B1"], Listed(sheet).Select(v => v.Range));
    }

    [Fact]
    public void Over_long_titles_and_messages_are_cut_to_Excels_limits()
    {
        var sheet = Load($"<dataValidation errorTitle=\"{new string('t', 40)}\" error=\"{new string('e', 300)}\" promptTitle=\"{new string('p', 33)}\" prompt=\"{new string('m', 256)}\" sqref=\"A1\"/>");

        var validation = sheet.Validations.Single().Validation;
        Assert.Equal(new string('t', 32), validation.ErrorTitle);
        Assert.Equal(new string('e', 255), validation.ErrorMessage);
        Assert.Equal(new string('p', 32), validation.InputTitle);
        Assert.Equal(new string('m', 255), validation.InputMessage);
    }

    [Fact]
    public void Rules_Excel_wrote_without_the_formulas_the_model_asks_for_are_read()
    {
        var sheet = Load("""
            <dataValidation type="whole" sqref="A1"/>
            <dataValidation type="whole" operator="between" sqref="A2"><formula1>1</formula1></dataValidation>
            <dataValidation type="list" operator="greaterThan" sqref="A3"><formula1>"a"</formula1><formula2>5</formula2></dataValidation>
            """);

        Assert.Equal(3, sheet.Validations.Count);
        Assert.Null(sheet["A1"].Validation!.Formula1);
        Assert.Null(sheet["A2"].Validation!.Formula2);
        Assert.Equal("=5", sheet["A3"].Validation!.Formula2);
    }

    [Fact]
    public void A_formula_SharpCell_cannot_read_is_kept_as_text()
    {
        var sheet = Load("""
            <dataValidation type="custom" sqref="A1"><formula1>SUM(A1:</formula1></dataValidation>
            <dataValidation type="list" sqref="A2"><formula1>[1]Book!$A$1:$A$3</formula1></dataValidation>
            """);

        Assert.Equal("=SUM(A1:", sheet["A1"].Validation!.Formula1);
        Assert.Equal("=[1]Book!$A$1:$A$3", sheet["A2"].Validation!.Formula1);
    }

    [Fact]
    public void Validations_in_the_extension_list_are_not_read()
    {
        var sheet = XlsxReader.Load(new TestXlsx().RawSheet("S", """
            <sheetData/>
            <extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
            <x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Other!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>A2</xm:sqref></x14:dataValidation>
            </x14:dataValidations></ext></extLst>
            """).Sheet("Other", "").Build())["S"];

        Assert.Empty(sheet.Validations);
    }
}
