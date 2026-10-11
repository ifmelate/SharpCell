using System;

namespace SharpCell;

/// <summary>What a data validation rule lets a cell hold, as Excel's Data Validation dialog names it under Allow.</summary>
public enum ValidationType
{
    /// <summary>Any value: the rule only shows its input message.</summary>
    Any,

    /// <summary>A whole number compared with the formulas.</summary>
    WholeNumber,

    /// <summary>A number compared with the formulas.</summary>
    Decimal,

    /// <summary>One of the items of a list, given as a quoted literal or a range; Excel shows them in a drop-down.</summary>
    List,

    /// <summary>A date compared with the formulas.</summary>
    Date,

    /// <summary>A time compared with the formulas.</summary>
    Time,

    /// <summary>Text whose length is compared with the formulas.</summary>
    TextLength,

    /// <summary>Any value for which the formula is true.</summary>
    Custom,
}

/// <summary>How a value is compared with a rule's formulas; lists and custom rules do not compare.</summary>
public enum ValidationOperator
{
    /// <summary>From the first formula to the second, both included.</summary>
    Between,

    /// <summary>Less than the first formula or greater than the second.</summary>
    NotBetween,

    /// <summary>Equal to the first formula.</summary>
    Equal,

    /// <summary>Not equal to the first formula.</summary>
    NotEqual,

    /// <summary>Greater than the first formula.</summary>
    GreaterThan,

    /// <summary>Less than the first formula.</summary>
    LessThan,

    /// <summary>Greater than or equal to the first formula.</summary>
    GreaterThanOrEqual,

    /// <summary>Less than or equal to the first formula.</summary>
    LessThanOrEqual,
}

/// <summary>What Excel does when a value breaks a rule.</summary>
public enum ValidationErrorStyle
{
    /// <summary>Refuses the value.</summary>
    Stop,

    /// <summary>Asks whether to keep the value.</summary>
    Warning,

    /// <summary>Tells the user and keeps the value.</summary>
    Information,
}

/// <summary>
/// A data validation rule: what the cells of a range accept and the messages Excel shows for them.
/// Set properties in an initializer, or start from <see cref="List"/>, <see cref="ListFrom"/> or
/// <see cref="Custom"/>. SharpCell keeps the rule with the cells and saves it; it does not check values.
/// </summary>
public sealed class DataValidation : IEquatable<DataValidation>
{
    private const int MaxTitleLength = 32;
    private const int MaxMessageLength = 255;
    private const int MaxListLength = 255;

    private readonly string? _formula1;
    private readonly string? _formula2;
    private readonly string? _inputTitle;
    private readonly string? _inputMessage;
    private readonly string? _errorTitle;
    private readonly string? _errorMessage;

    /// <summary>A rule allowing any value; set properties in an initializer.</summary>
    public DataValidation()
    {
    }

    /// <summary>What the cells may hold; <see cref="ValidationType.Any"/> unless set.</summary>
    public ValidationType Type { get; init; }

    /// <summary>
    /// How a value is compared with <see cref="Formula1"/> and <see cref="Formula2"/>;
    /// <see cref="ValidationOperator.Between"/> unless set. Lists and custom rules ignore it.
    /// </summary>
    public ValidationOperator Operator { get; init; }

    /// <summary>
    /// The first formula, with a leading <c>=</c> (added if missing) as <see cref="Cell.Formula"/> takes it, or null:
    /// the low end or the value compared with, the list (<c>="Yes,No"</c> or <c>=$H$1:$H$5</c>) or the custom
    /// condition (<c>=A1&gt;0</c>). Relative references are relative to the first cell of the rule's first range.
    /// It is parsed when the rule is added to a sheet.
    /// </summary>
    public string? Formula1
    {
        get => _formula1;
        init => _formula1 = WithEquals(value);
    }

    /// <summary>
    /// The second formula, the high end for <see cref="ValidationOperator.Between"/> and
    /// <see cref="ValidationOperator.NotBetween"/>, written like <see cref="Formula1"/>; null for none.
    /// </summary>
    public string? Formula2
    {
        get => _formula2;
        init => _formula2 = WithEquals(value);
    }

    /// <summary>Whether an empty cell is accepted; true unless set.</summary>
    public bool AllowBlank { get; init; } = true;

    /// <summary>Whether Excel shows the drop-down of a list rule; true unless set.</summary>
    public bool ShowDropDown { get; init; } = true;

    /// <summary>Whether Excel shows the input message when a cell of the rule is selected; true unless set.</summary>
    public bool ShowInputMessage { get; init; } = true;

    /// <summary>The title of the input message, at most 32 characters, or null for none.</summary>
    /// <exception cref="ArgumentException">The title is longer than 32 characters.</exception>
    public string? InputTitle
    {
        get => _inputTitle;
        init => _inputTitle = Checked(value, MaxTitleLength, nameof(InputTitle));
    }

    /// <summary>The input message, at most 255 characters, or null for none.</summary>
    /// <exception cref="ArgumentException">The message is longer than 255 characters.</exception>
    public string? InputMessage
    {
        get => _inputMessage;
        init => _inputMessage = Checked(value, MaxMessageLength, nameof(InputMessage));
    }

    /// <summary>Whether Excel shows the error message when a value breaks the rule; true unless set.</summary>
    public bool ShowErrorMessage { get; init; } = true;

    /// <summary>What Excel does with a value that breaks the rule; <see cref="ValidationErrorStyle.Stop"/> unless set.</summary>
    public ValidationErrorStyle ErrorStyle { get; init; }

    /// <summary>The title of the error message, at most 32 characters, or null for Excel's own.</summary>
    /// <exception cref="ArgumentException">The title is longer than 32 characters.</exception>
    public string? ErrorTitle
    {
        get => _errorTitle;
        init => _errorTitle = Checked(value, MaxTitleLength, nameof(ErrorTitle));
    }

    /// <summary>The error message, at most 255 characters, or null for Excel's own.</summary>
    /// <exception cref="ArgumentException">The message is longer than 255 characters.</exception>
    public string? ErrorMessage
    {
        get => _errorMessage;
        init => _errorMessage = Checked(value, MaxMessageLength, nameof(ErrorMessage));
    }

    /// <summary>
    /// A drop-down list of the given items, kept in the rule as one quoted literal:
    /// <c>List("Yes", "No")</c> has the formula <c>="Yes,No"</c>.
    /// </summary>
    /// <param name="items">The items, in order. An item cannot hold a comma, which separates them.</param>
    /// <exception cref="ArgumentNullException">The items or one of them is null.</exception>
    /// <exception cref="ArgumentException">There are no items, an item holds a comma, or the items with their commas
    /// are longer than 255 characters, Excel's limit.</exception>
    public static DataValidation List(params string[] items)
    {
        ArgumentNullException.ThrowIfNull(items);
        if (items.Length == 0)
            throw new ArgumentException("A list needs at least one item.", nameof(items));
        foreach (var item in items)
        {
            ArgumentNullException.ThrowIfNull(item, nameof(items));
            if (item.Contains(','))
                throw new ArgumentException($"The item '{item}' holds a comma, which separates list items.", nameof(items));
        }

        var list = string.Join(",", items);
        if (list.Length > MaxListLength)
            throw new ArgumentException($"The items with their commas are longer than {MaxListLength} characters; put them in a range and use ListFrom.", nameof(items));
        return new DataValidation { Type = ValidationType.List, Formula1 = "=\"" + list.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"" };
    }

    /// <summary>A drop-down list of the values of a range, such as <c>ListFrom("$H$1:$H$5")</c> or a defined name.</summary>
    /// <param name="range">The range or name, with or without a leading <c>=</c>.</param>
    /// <exception cref="ArgumentException">The range is null or empty.</exception>
    public static DataValidation ListFrom(string range)
    {
        ArgumentException.ThrowIfNullOrEmpty(range);
        return new DataValidation { Type = ValidationType.List, Formula1 = range };
    }

    /// <summary>A rule accepting a value when the formula is true, such as <c>Custom("=A1&gt;0")</c>.</summary>
    /// <param name="formula">The condition, relative to the first cell of the rule's first range.</param>
    /// <exception cref="ArgumentException">The formula is null or empty.</exception>
    public static DataValidation Custom(string formula)
    {
        ArgumentException.ThrowIfNullOrEmpty(formula);
        return new DataValidation { Type = ValidationType.Custom, Formula1 = formula };
    }

    /// <summary>A copy with other formula texts, for a rule whose first cell moved.</summary>
    internal DataValidation WithFormulas(string? formula1, string? formula2) => new()
    {
        Type = Type,
        Operator = Operator,
        Formula1 = formula1,
        Formula2 = formula2,
        AllowBlank = AllowBlank,
        ShowDropDown = ShowDropDown,
        ShowInputMessage = ShowInputMessage,
        InputTitle = InputTitle,
        InputMessage = InputMessage,
        ShowErrorMessage = ShowErrorMessage,
        ErrorStyle = ErrorStyle,
        ErrorTitle = ErrorTitle,
        ErrorMessage = ErrorMessage,
    };

    /// <summary>Whether the other rule has the same settings; texts are compared ordinally.</summary>
    public bool Equals(DataValidation? other) =>
        other is not null
        && Type == other.Type
        && Operator == other.Operator
        && string.Equals(Formula1, other.Formula1, StringComparison.Ordinal)
        && string.Equals(Formula2, other.Formula2, StringComparison.Ordinal)
        && AllowBlank == other.AllowBlank
        && ShowDropDown == other.ShowDropDown
        && ShowInputMessage == other.ShowInputMessage
        && string.Equals(InputTitle, other.InputTitle, StringComparison.Ordinal)
        && string.Equals(InputMessage, other.InputMessage, StringComparison.Ordinal)
        && ShowErrorMessage == other.ShowErrorMessage
        && ErrorStyle == other.ErrorStyle
        && string.Equals(ErrorTitle, other.ErrorTitle, StringComparison.Ordinal)
        && string.Equals(ErrorMessage, other.ErrorMessage, StringComparison.Ordinal);

    /// <summary>Whether <paramref name="obj"/> is a rule with the same settings.</summary>
    public override bool Equals(object? obj) => Equals(obj as DataValidation);

    /// <summary>A hash code consistent with <see cref="Equals(DataValidation)"/>.</summary>
    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(Type);
        hash.Add(Operator);
        hash.Add(Formula1);
        hash.Add(Formula2);
        hash.Add(AllowBlank);
        hash.Add(ShowDropDown);
        hash.Add(ShowInputMessage);
        hash.Add(InputTitle);
        hash.Add(InputMessage);
        hash.Add(ShowErrorMessage);
        hash.Add(ErrorStyle);
        hash.Add(ErrorTitle);
        hash.Add(ErrorMessage);
        return hash.ToHashCode();
    }

    /// <summary>The type, the operator of a comparing type, and the formulas, such as <c>WholeNumber Between =1 =10</c>.</summary>
    public override string ToString()
    {
        var text = Type is ValidationType.Any or ValidationType.List or ValidationType.Custom ? Type.ToString() : $"{Type} {Operator}";
        if (Formula1 is not null)
            text += " " + Formula1;
        if (Formula2 is not null)
            text += " " + Formula2;
        return text;
    }

    private static string? WithEquals(string? formula) => formula is null || formula.StartsWith('=') ? formula : "=" + formula;

    private static string? Checked(string? text, int max, string property)
    {
        if (text?.Length > max)
            throw new ArgumentException($"{property} is at most {max} characters.", "value");
        return text;
    }
}
