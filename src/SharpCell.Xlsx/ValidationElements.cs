using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Xlsx;

/// <summary>
/// A sheet's data validation rules in its &lt;dataValidations&gt; element: one &lt;dataValidation&gt; per rule,
/// its ranges in sqref and its formulas in file syntax, relative to the first cell of sqref. Rules in
/// the extension list (x14:dataValidations) are not read; the file keeps them.
/// </summary>
internal static class ValidationElements
{
    // Excel's limits, which the model checks; a file that breaks them has its texts cut.
    private const int MaxTitleLength = 32;
    private const int MaxMessageLength = 255;

    /// <summary>
    /// Adds the element's rules to the sheet as the file has them. Validation only decorates cells,
    /// so a rule with a damaged range or an unknown setting is skipped, never a reason to refuse the file.
    /// </summary>
    public static void Read(XElement element, Worksheet sheet)
    {
        foreach (var item in element.Elements(element.Name.Namespace + "dataValidation"))
        {
            if (Rule(item) is { } rule)
                sheet.AddValidationRule(rule.Areas, rule.Validation);
        }
    }

    private static (List<Area> Areas, DataValidation Validation)? Rule(XElement element)
    {
        if (Areas((string?)element.Attribute("sqref")) is not { } areas
            || TypeOf((string?)element.Attribute("type") ?? "none") is not { } type
            || OperatorOf((string?)element.Attribute("operator") ?? "between") is not { } op
            || ErrorStyleOf((string?)element.Attribute("errorStyle") ?? "stop") is not { } style)
            return null;

        // The schema's flags are off unless set; showDropDown="1" hides the drop-down.
        var origin = new CellAddress(areas[0].FirstRow, areas[0].FirstColumn);
        var ns = element.Name.Namespace;
        return (areas, new DataValidation
        {
            Type = type,
            Operator = op,
            Formula1 = Formula(element.Element(ns + "formula1"), origin),
            Formula2 = Formula(element.Element(ns + "formula2"), origin),
            AllowBlank = CellMetadata.IsTrue((string?)element.Attribute("allowBlank")),
            ShowDropDown = !CellMetadata.IsTrue((string?)element.Attribute("showDropDown")),
            ShowInputMessage = CellMetadata.IsTrue((string?)element.Attribute("showInputMessage")),
            InputTitle = Cut((string?)element.Attribute("promptTitle"), MaxTitleLength),
            InputMessage = Cut((string?)element.Attribute("prompt"), MaxMessageLength),
            ShowErrorMessage = CellMetadata.IsTrue((string?)element.Attribute("showErrorMessage")),
            ErrorStyle = style,
            ErrorTitle = Cut((string?)element.Attribute("errorTitle"), MaxTitleLength),
            ErrorMessage = Cut((string?)element.Attribute("error"), MaxMessageLength),
        });
    }

    // The ranges of a sqref, or null when it has none or one is not a range.
    private static List<Area>? Areas(string? sqref)
    {
        var areas = new List<Area>();
        foreach (var part in (sqref ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (!ReferenceSyntax.TryParseA1Area(part.Replace("$", "", StringComparison.Ordinal), new CellAddress(1, 1), out var parsed))
                return null;
            areas.Add(Area.Resolve(parsed, new CellAddress(1, 1)));
        }

        return areas.Count > 0 ? areas : null;
    }

    // The formula as the model takes it: printed again without the file's prefixes when SharpCell can
    // read it, the file's text otherwise (the sheet keeps such a formula as text).
    private static string? Formula(XElement? element, CellAddress origin)
    {
        if (element?.Value is not { Length: > 0 } text)
            return null;
        if (FormulaText.ReferencesOtherWorkbook(text))
            return "=" + text;
        try
        {
            return "=" + FormulaPrinter.Print(FormulaParser.Parse("=" + text, origin), origin);
        }
        catch (FormulaParseException)
        {
            return "=" + text;
        }
    }

    private static string? Cut(string? text, int max)
    {
        if (text is null || text.Length <= max)
            return text;
        return text[..(char.IsHighSurrogate(text[max - 1]) ? max - 1 : max)];
    }

    /// <summary>The &lt;dataValidations&gt; element of the sheet's rules, with the file element's own attributes; null when there are none.</summary>
    public static XElement? Element(XNamespace ns, Worksheet sheet, XElement? source)
    {
        var rules = sheet.ValidationRules;
        if (rules.Count == 0)
            return null;
        var element = new XElement(ns + "dataValidations", source?.Attributes());
        element.SetAttributeValue("count", rules.Count);
        foreach (var rule in rules)
        {
            var validation = rule.Validation;
            element.Add(new XElement(ns + "dataValidation",
                validation.Type == ValidationType.Any ? null : new XAttribute("type", TypeText(validation.Type)),
                validation.Operator == ValidationOperator.Between ? null : new XAttribute("operator", OperatorText(validation.Operator)),
                validation.AllowBlank ? new XAttribute("allowBlank", "1") : null,
                validation.ShowDropDown ? null : new XAttribute("showDropDown", "1"),
                validation.ShowInputMessage ? new XAttribute("showInputMessage", "1") : null,
                validation.ShowErrorMessage ? new XAttribute("showErrorMessage", "1") : null,
                validation.ErrorStyle == ValidationErrorStyle.Stop ? null : new XAttribute("errorStyle", ErrorStyleText(validation.ErrorStyle)),
                validation.ErrorTitle is { } errorTitle ? new XAttribute("errorTitle", errorTitle) : null,
                validation.ErrorMessage is { } error ? new XAttribute("error", error) : null,
                validation.InputTitle is { } promptTitle ? new XAttribute("promptTitle", promptTitle) : null,
                validation.InputMessage is { } prompt ? new XAttribute("prompt", prompt) : null,
                new XAttribute("sqref", string.Join(' ', rule.Areas.Select(WorksheetParts.Format))),
                FormulaElement(ns + "formula1", rule.F1, rule.Origin, sheet),
                FormulaElement(ns + "formula2", rule.F2, rule.Origin, sheet)));
        }

        return element;
    }

    // A formula SharpCell could not read is written as the file had it.
    private static XElement? FormulaElement(XName name, FormulaNode? formula, CellAddress origin, Worksheet sheet) =>
        formula is null
            ? null
            : new XElement(name, FormulaPrinter.PrintForFile(formula, origin, at => sheet.Workbook.TableAt(sheet, at.Row, at.Column)?.Name,
                function => FunctionValues.IsFunction(sheet.Workbook, sheet, function)));

    private static ValidationType? TypeOf(string text) => text switch
    {
        "none" => ValidationType.Any,
        "whole" => ValidationType.WholeNumber,
        "decimal" => ValidationType.Decimal,
        "list" => ValidationType.List,
        "date" => ValidationType.Date,
        "time" => ValidationType.Time,
        "textLength" => ValidationType.TextLength,
        "custom" => ValidationType.Custom,
        _ => null,
    };

    private static string TypeText(ValidationType type) => type switch
    {
        ValidationType.WholeNumber => "whole",
        ValidationType.Decimal => "decimal",
        ValidationType.List => "list",
        ValidationType.Date => "date",
        ValidationType.Time => "time",
        ValidationType.TextLength => "textLength",
        ValidationType.Custom => "custom",
        _ => "none",
    };

    private static ValidationOperator? OperatorOf(string text) => text switch
    {
        "between" => ValidationOperator.Between,
        "notBetween" => ValidationOperator.NotBetween,
        "equal" => ValidationOperator.Equal,
        "notEqual" => ValidationOperator.NotEqual,
        "greaterThan" => ValidationOperator.GreaterThan,
        "lessThan" => ValidationOperator.LessThan,
        "greaterThanOrEqual" => ValidationOperator.GreaterThanOrEqual,
        "lessThanOrEqual" => ValidationOperator.LessThanOrEqual,
        _ => null,
    };

    private static string OperatorText(ValidationOperator op) => op switch
    {
        ValidationOperator.NotBetween => "notBetween",
        ValidationOperator.Equal => "equal",
        ValidationOperator.NotEqual => "notEqual",
        ValidationOperator.GreaterThan => "greaterThan",
        ValidationOperator.LessThan => "lessThan",
        ValidationOperator.GreaterThanOrEqual => "greaterThanOrEqual",
        ValidationOperator.LessThanOrEqual => "lessThanOrEqual",
        _ => "between",
    };

    private static ValidationErrorStyle? ErrorStyleOf(string text) => text switch
    {
        "stop" => ValidationErrorStyle.Stop,
        "warning" => ValidationErrorStyle.Warning,
        "information" => ValidationErrorStyle.Information,
        _ => null,
    };

    private static string ErrorStyleText(ValidationErrorStyle style) => style == ValidationErrorStyle.Warning ? "warning" : "information";
}
