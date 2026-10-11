using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>Data validation rules are saved in the sheet's &lt;dataValidations&gt; element, written from the model when they changed.</summary>
public class XlsxWriterValidationTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    private const string Extension = """
        <extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
        <x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Other!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>G4</xm:sqref></x14:dataValidation>
        </x14:dataValidations></ext></extLst>
        """;

    private static TestXlsx Validated(string tail = "") => new TestXlsx().RawSheet("S", $"""
        <sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="5"><c r="A5"><v>5</v></c></row></sheetData>
        <dataValidations disablePrompts="1" count="4">
        <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B2:B6"><formula1>"Yes,No"</formula1></dataValidation>
        <dataValidation type="whole" allowBlank="1" showErrorMessage="1" errorTitle="Count" error="1 to 10" sqref="C2:C6 E2"><formula1>1</formula1><formula2>$A$5</formula2></dataValidation>
        <dataValidation type="custom" imeMode="hiragana" sqref="D5:D10"><formula1>D5&gt;A5</formula1></dataValidation>
        <dataValidation type="list" sqref="F1"><formula1>Other!$A$1:$A$3</formula1></dataValidation>
        </dataValidations>
        <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        {tail}
        """).Sheet("Other", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>");

    private static TestXlsx Unvalidated() => new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>").Sheet("Other", "");

    private static readonly DataValidation PositiveWholeNumber = new()
    {
        Type = ValidationType.WholeNumber,
        Operator = ValidationOperator.GreaterThan,
        Formula1 = "=0",
        ShowDropDown = false,
        InputTitle = "Count",
        InputMessage = "How many?",
        ErrorStyle = ValidationErrorStyle.Information,
        ErrorTitle = "Not a count",
        ErrorMessage = "A <positive> & \"whole\" number.",
    };

    public static TheoryData<string, string> Changes => new()
    {
        { "validated", "list added" }, { "validated", "whole number added" }, { "validated", "one changed" },
        { "validated", "part cleared" }, { "validated", "all cleared" }, { "validated", "rows inserted" },
        { "validated", "rows deleted" }, { "validated", "columns deleted" }, { "validated", "rows inserted on another sheet" },
        { "unvalidated", "list added" }, { "unvalidated", "whole number added" },
    };

    private static void Change(Worksheet s, string change)
    {
        switch (change)
        {
            case "list added": s.AddValidation("G1:G3", DataValidation.List("a", "b")); break;
            case "whole number added": s.AddValidation("H1", PositiveWholeNumber); break;
            case "one changed": s.AddValidation("B2:B6", DataValidation.List("Yes", "No", "Maybe")); break;
            case "part cleared": s.ClearValidation("C2:D6"); break;
            case "all cleared": s.ClearValidation("A1:Z100"); break;
            case "rows inserted": s.InsertRows(3, 2); break;
            case "rows deleted": s.DeleteRows(1, 2); break;
            case "columns deleted": s.DeleteColumns(3); break;
            case "rows inserted on another sheet": s.Workbook["Other"].InsertRows(1); break;
        }
    }

    private static List<(string Range, DataValidation Validation)> Listed(Worksheet sheet) =>
        [.. sheet.Validations.Select(v => (v.Range.Address, v.Validation))];

    [Theory]
    [MemberData(nameof(Changes))]
    public void A_change_is_saved_and_reads_back(string file, string change)
    {
        var original = (file == "validated" ? Validated() : Unvalidated()).Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        Change(workbook["S"], change);
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var reloaded = XlsxWriterTests.Reload(saved);
        Assert.Equal(Listed(workbook["S"]), Listed(reloaded["S"]));
    }

    [Fact]
    public void A_workbook_built_in_code_saves_its_rules()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Report");
        sheet["A1"].Value = "Answer";
        sheet.AddValidation("A2:A10", DataValidation.List("Yes", "No"));
        sheet.AddValidation("B2:B10", PositiveWholeNumber);
        sheet.AddValidation("C2:C10", DataValidation.Custom("=ISOWEEKNUM(C2)>B2"));
        sheet.AddValidation("D2", new DataValidation { Type = ValidationType.Date, Formula1 = "=DATE(2026,1,1)", Formula2 = "=TODAY()" });
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        var part = XlsxWriterTests.Part(saved, Sheet1);
        Assert.Contains("<dataValidations count=\"4\">", part);
        Assert.Contains("<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A2:A10\">"
            + "<formula1>\"Yes,No\"</formula1></dataValidation>", part);
        Assert.Contains("<dataValidation type=\"whole\" operator=\"greaterThan\" allowBlank=\"1\" showDropDown=\"1\" showInputMessage=\"1\" "
            + "showErrorMessage=\"1\" errorStyle=\"information\" errorTitle=\"Not a count\" error=\"A &lt;positive&gt; &amp; &quot;whole&quot; number.\" "
            + "promptTitle=\"Count\" prompt=\"How many?\" sqref=\"B2:B10\"><formula1>0</formula1></dataValidation>", part);
        Assert.Contains("<formula1>_xlfn.ISOWEEKNUM(C2)&gt;B2</formula1>", part);
        Assert.Contains("<formula1>DATE(2026,1,1)</formula1><formula2>TODAY()</formula2>", part);
        var reloaded = XlsxWriterTests.Reload(saved);
        Assert.Equal(Listed(sheet), Listed(reloaded["Report"]));
        Assert.Equal("Answer", reloaded["Report"]["A1"].Value.AsText());
    }

    [Fact]
    public void Control_characters_and_escape_like_text_in_messages_round_trip()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet.AddValidation("A1", new DataValidation
        {
            InputTitle = "_x0041_",
            InputMessage = "line\r\nbreak\ttab",
            ErrorTitle = "a\u0001b",
            ErrorMessage = "\u0002 control _x005F_",
        });
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        var part = XlsxWriterTests.Part(saved, Sheet1);
        Assert.Contains("errorTitle=\"a_x0001_b\"", part);
        Assert.Contains("promptTitle=\"_x005F_x0041_\"", part);
        Assert.Equal(Listed(sheet), Listed(XlsxWriterTests.Reload(saved)["S"]));
    }

    [Fact]
    public void Changing_a_rule_rewrites_only_the_rules()
    {
        var original = Validated().Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook["S"].AddValidation("B2:B6", DataValidation.List("Yes", "No", "Maybe"));
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        Assert.Equal([Sheet1], XlsxWriterHyperlinkTests.ChangedParts(original, saved));
        Assert.True(XNode.DeepEquals(WithoutValidations(original), WithoutValidations(saved)));
        var part = XlsxWriterTests.Part(saved, Sheet1);
        Assert.Contains("<dataValidations disablePrompts=\"1\" count=\"4\">", part);
        Assert.Contains("<formula1>\"Yes,No,Maybe\"</formula1>", part);
    }

    [Fact]
    public void A_workbook_with_rules_saved_without_changes_is_the_source_bytes()
    {
        var original = Validated(Extension).Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();

        Assert.Equal(original, XlsxWriterTests.Save(workbook));
    }

    [Fact]
    public void Validations_in_the_extension_list_survive_a_change_to_the_rules()
    {
        var original = Validated(Extension).Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook["S"].ClearValidation("A1:Z100");
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var part = XlsxWriterTests.Part(saved, Sheet1);
        Assert.DoesNotContain("<dataValidations", part);
        Assert.Contains("<xm:f>Other!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>G4</xm:sqref>", part);
    }

    [Fact]
    public void Validations_in_the_extension_list_still_follow_deleted_rows()
    {
        var original = Validated(Extension).Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook["S"].DeleteRows(1, 2);
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var part = XlsxWriterTests.Part(saved, Sheet1);
        Assert.Contains("<xm:sqref>G2</xm:sqref>", part);
        Assert.Contains("sqref=\"D3:D8\"><formula1>D3&gt;A3</formula1>", part);
    }

    [Fact]
    public void Rules_follow_deleted_rows_once()
    {
        var original = new TestXlsx().RawSheet("S", """
            <sheetData/><dataValidations count="2">
            <dataValidation type="custom" sqref="A5:A10"><formula1>A5&gt;0</formula1></dataValidation>
            <dataValidation type="whole" operator="lessThan" sqref="C5"><formula1>$B$10</formula1></dataValidation>
            </dataValidations>
            """).Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook["S"].DeleteRows(1, 2);
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var part = XlsxWriterTests.Part(saved, Sheet1);
        Assert.Contains("sqref=\"A3:A8\"><formula1>A3&gt;0</formula1>", part);
        Assert.Contains("sqref=\"C3\"><formula1>$B$8</formula1>", part);

        // The rule loses its first cells: its formula moves to the new first cell.
        workbook["S"].ClearValidation("A3:A4");
        workbook.Recalculate();
        part = XlsxWriterTests.Part(XlsxWriterTests.Save(workbook), Sheet1);
        Assert.Contains("sqref=\"A5:A8\"><formula1>A5&gt;0</formula1>", part);
    }

    [Fact]
    public void Rules_the_model_would_not_take_and_unreadable_formulas_are_written_as_read()
    {
        var original = new TestXlsx().RawSheet("S", """
            <sheetData/><dataValidations count="4">
            <dataValidation type="whole" sqref="A1"/>
            <dataValidation type="list" operator="greaterThan" sqref="A2"><formula1>"a"</formula1><formula2>5</formula2></dataValidation>
            <dataValidation type="custom" sqref="A3"><formula1>SUM(A3:</formula1></dataValidation>
            <dataValidation type="list" sqref="A4"><formula1>[1]Book!$A$1:$A$3</formula1></dataValidation>
            </dataValidations>
            """).Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook["S"].AddValidation("B1", DataValidation.List("x"));
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        var part = XlsxWriterTests.Part(saved, Sheet1);
        Assert.Contains("<dataValidation type=\"whole\" sqref=\"A1\" />", part);
        Assert.Contains("<dataValidation type=\"list\" operator=\"greaterThan\" sqref=\"A2\"><formula1>\"a\"</formula1><formula2>5</formula2>", part);
        Assert.Contains("<formula1>SUM(A3:</formula1>", part);
        Assert.Contains("<formula1>[1]Book!$A$1:$A$3</formula1>", part);
        Assert.Equal(Listed(workbook["S"]), Listed(XlsxWriterTests.Reload(saved)["S"]));
    }

    [Fact]
    public void A_damaged_rule_is_not_left_behind_when_rows_move()
    {
        var original = new TestXlsx().RawSheet("S", """
            <sheetData/><dataValidations count="1"><dataValidation type="money" sqref="A5"><formula1>1</formula1></dataValidation></dataValidations>
            """).Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook["S"].DeleteRows(1);
        workbook.Recalculate();

        var part = XlsxWriterTests.Part(XlsxWriterTests.Save(workbook), Sheet1);

        Assert.DoesNotContain("dataValidation", part);
    }

    [Fact]
    public void A_rule_over_several_ranges_is_saved_as_one_element()
    {
        var workbook = XlsxWriterTests.Load(Validated());
        workbook["S"].InsertColumns(4);
        workbook.Recalculate();

        var part = XlsxWriterTests.Part(XlsxWriterTests.Save(workbook), Sheet1);

        Assert.Single(Regex.Matches(part, "sqref=\"C2:C6 F2\""));
    }

    private static XDocument WithoutValidations(byte[] file)
    {
        var document = XDocument.Parse(XlsxWriterTests.Part(file, Sheet1));
        document.Root!.Elements().Where(e => e.Name.LocalName == "dataValidations").Remove();
        return document;
    }
}
