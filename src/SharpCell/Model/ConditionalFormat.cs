using System;

namespace SharpCell;

/// <summary>What a conditional format looks at, as Excel's New Formatting Rule dialog offers it.</summary>
public enum ConditionalFormatType
{
    /// <summary>The cell's value compared with formulas: Highlight Cells Rules &gt; Greater Than and the others.</summary>
    CellValue,

    /// <summary>Cells for which a formula is true: Use a formula to determine which cells to format.</summary>
    Expression,

    /// <summary>The highest values of the ranges: Top/Bottom Rules &gt; Top 10 Items or Top 10%.</summary>
    Top,

    /// <summary>The lowest values of the ranges.</summary>
    Bottom,

    /// <summary>Values above the average of the ranges.</summary>
    AboveAverage,

    /// <summary>Values below the average of the ranges.</summary>
    BelowAverage,

    /// <summary>Values that appear more than once in the ranges.</summary>
    DuplicateValues,

    /// <summary>Values that appear once in the ranges.</summary>
    UniqueValues,

    /// <summary>Text containing <see cref="ConditionalFormat.Text"/>, ignoring case.</summary>
    ContainsText,

    /// <summary>Text not containing <see cref="ConditionalFormat.Text"/>.</summary>
    NotContainsText,

    /// <summary>Text beginning with <see cref="ConditionalFormat.Text"/>.</summary>
    BeginsWith,

    /// <summary>Text ending with <see cref="ConditionalFormat.Text"/>.</summary>
    EndsWith,

    /// <summary>Empty cells, or cells holding only spaces.</summary>
    Blanks,

    /// <summary>Cells that are not blank.</summary>
    NoBlanks,

    /// <summary>Cells holding an error.</summary>
    Errors,

    /// <summary>Cells not holding an error.</summary>
    NoErrors,

    /// <summary>Dates in a period around today: A Date Occurring.</summary>
    TimePeriod,

    /// <summary>A colour scale, read from a file; every number gets a colour of its own, which SharpCell does not work out.</summary>
    ColorScale,

    /// <summary>Data bars, read from a file.</summary>
    DataBar,

    /// <summary>An icon set, read from a file.</summary>
    IconSet,
}

/// <summary>How a <see cref="ConditionalFormatType.CellValue"/> rule compares the cell's value with its formulas.</summary>
public enum ConditionalOperator
{
    /// <summary>From the first formula to the second, both included.</summary>
    Between,

    /// <summary>Less than the first formula or greater than the second.</summary>
    NotBetween,

    /// <summary>Equal to the first formula; text ignoring case.</summary>
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

/// <summary>The periods of Excel's A Date Occurring rule, counted from today.</summary>
public enum TimePeriod
{
    /// <summary>Yesterday.</summary>
    Yesterday,

    /// <summary>Today.</summary>
    Today,

    /// <summary>Tomorrow.</summary>
    Tomorrow,

    /// <summary>Today and the six days before.</summary>
    Last7Days,

    /// <summary>The week before this one, Sunday to Saturday.</summary>
    LastWeek,

    /// <summary>This week, Sunday to Saturday.</summary>
    ThisWeek,

    /// <summary>The week after this one.</summary>
    NextWeek,

    /// <summary>The month before this one.</summary>
    LastMonth,

    /// <summary>This month.</summary>
    ThisMonth,

    /// <summary>The month after this one.</summary>
    NextMonth,
}

/// <summary>
/// A conditional format: which cells of its ranges it formats, and how. Make one with a factory
/// such as <see cref="CellValue"/>, <see cref="Expression"/> or <see cref="Top"/>, change it with
/// <c>with</c> (<c>rule with { StopIfTrue = true }</c>), and add it with
/// <see cref="Worksheet.AddConditionalFormat"/>. SharpCell works out which cells a rule applies to
/// (<see cref="Cell.ConditionalFormats"/>) and how a cell then looks (<see cref="Cell.DisplayStyle"/>).
/// </summary>
public sealed record ConditionalFormat
{
    private ConditionalFormat(ConditionalFormatType type, ConditionalStyle? style)
    {
        Type = type;
        Style = style;
    }

    /// <summary>What the rule looks at.</summary>
    public ConditionalFormatType Type { get; }

    /// <summary>How a <see cref="ConditionalFormatType.CellValue"/> rule compares; <see cref="ConditionalOperator.Between"/> for other rules.</summary>
    public ConditionalOperator Operator { get; private init; }

    /// <summary>
    /// The formula of a <see cref="ConditionalFormatType.CellValue"/> or <see cref="ConditionalFormatType.Expression"/>
    /// rule, starting with <c>=</c>, relative to the first cell of the rule's first range; null for other rules.
    /// </summary>
    public string? Formula1 { get; private init; }

    /// <summary>The second formula of a between rule; null for others.</summary>
    public string? Formula2 { get; private init; }

    /// <summary>The text a text rule looks for; null for others.</summary>
    public string? Text { get; private init; }

    /// <summary>How many values (or what percentage of them) a top or bottom rule formats; 0 for others.</summary>
    public int Rank { get; private init; }

    /// <summary>Whether <see cref="Rank"/> is a percentage.</summary>
    public bool Percent { get; private init; }

    /// <summary>Whether an above or below average rule also formats values equal to the average.</summary>
    public bool OrEqual { get; private init; }

    /// <summary>How many standard deviations above or below the average a value must be; 0 for the average itself.</summary>
    public int StandardDeviations { get; private init; }

    /// <summary>The period of a <see cref="ConditionalFormatType.TimePeriod"/> rule.</summary>
    public TimePeriod Period { get; private init; }

    /// <summary>How matching cells look; null for colour scales, data bars and icon sets, which colour each cell themselves.</summary>
    public ConditionalStyle? Style { get; init; }

    /// <summary>
    /// Whether lower rules are skipped for a cell this rule applies to, as Excel's Stop If True does.
    /// </summary>
    public bool StopIfTrue { get; init; }

    /// <summary>Cells whose value compares with the formulas, such as greater than <c>=100</c> or between <c>=$B$1</c> and <c>=$B$2</c>.</summary>
    /// <param name="comparison">How the value compares.</param>
    /// <param name="formula1">The formula compared with, starting with <c>=</c>, relative to the range's first cell.</param>
    /// <param name="style">How matching cells look.</param>
    /// <param name="formula2">The second formula, for <see cref="ConditionalOperator.Between"/> and <see cref="ConditionalOperator.NotBetween"/>.</param>
    /// <exception cref="ArgumentException">A formula is missing or empty, or a second formula is given to a comparison that takes one.</exception>
    public static ConditionalFormat CellValue(ConditionalOperator comparison, string formula1, ConditionalStyle style, string? formula2 = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(formula1);
        ArgumentNullException.ThrowIfNull(style);
        var between = comparison is ConditionalOperator.Between or ConditionalOperator.NotBetween;
        if (between != (formula2 is not null))
            throw new ArgumentException(between ? $"{comparison} needs a second formula." : $"{comparison} takes one formula.", nameof(formula2));
        if (formula2 is { Length: 0 })
            throw new ArgumentException("The second formula is empty.", nameof(formula2));
        return new ConditionalFormat(ConditionalFormatType.CellValue, style) { Operator = comparison, Formula1 = formula1, Formula2 = formula2 };
    }

    /// <summary>Cells for which a formula is true, such as <c>=$C2&gt;$B2</c>, written for the range's first cell.</summary>
    /// <param name="formula">The formula, starting with <c>=</c>.</param>
    /// <param name="style">How matching cells look.</param>
    /// <exception cref="ArgumentException">The formula is empty.</exception>
    public static ConditionalFormat Expression(string formula, ConditionalStyle style)
    {
        ArgumentException.ThrowIfNullOrEmpty(formula);
        ArgumentNullException.ThrowIfNull(style);
        return new ConditionalFormat(ConditionalFormatType.Expression, style) { Formula1 = formula };
    }

    /// <summary>The <paramref name="rank"/> highest numbers of the ranges, or the highest <paramref name="rank"/> percent of them.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The rank is outside 1 to 1000, or a percentage outside 1 to 100.</exception>
    public static ConditionalFormat Top(int rank, ConditionalStyle style, bool percent = false) => Ranked(ConditionalFormatType.Top, rank, style, percent);

    /// <summary>The <paramref name="rank"/> lowest numbers of the ranges, or the lowest <paramref name="rank"/> percent of them.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The rank is outside 1 to 1000, or a percentage outside 1 to 100.</exception>
    public static ConditionalFormat Bottom(int rank, ConditionalStyle style, bool percent = false) => Ranked(ConditionalFormatType.Bottom, rank, style, percent);

    /// <summary>Numbers above the average of the ranges, or by <paramref name="standardDeviations"/> standard deviations above it.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The standard deviations are outside 0 to 3, as in Excel.</exception>
    public static ConditionalFormat AboveAverage(ConditionalStyle style, bool orEqual = false, int standardDeviations = 0) =>
        Average(ConditionalFormatType.AboveAverage, style, orEqual, standardDeviations);

    /// <summary>Numbers below the average of the ranges, or by <paramref name="standardDeviations"/> standard deviations below it.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The standard deviations are outside 0 to 3, as in Excel.</exception>
    public static ConditionalFormat BelowAverage(ConditionalStyle style, bool orEqual = false, int standardDeviations = 0) =>
        Average(ConditionalFormatType.BelowAverage, style, orEqual, standardDeviations);

    /// <summary>Values that appear more than once in the ranges; text ignoring case.</summary>
    public static ConditionalFormat DuplicateValues(ConditionalStyle style) => Plain(ConditionalFormatType.DuplicateValues, style);

    /// <summary>Values that appear once in the ranges.</summary>
    public static ConditionalFormat UniqueValues(ConditionalStyle style) => Plain(ConditionalFormatType.UniqueValues, style);

    /// <summary>Cells whose text contains <paramref name="text"/>, ignoring case.</summary>
    /// <exception cref="ArgumentException">The text is empty.</exception>
    public static ConditionalFormat ContainsText(string text, ConditionalStyle style) => Textual(ConditionalFormatType.ContainsText, text, style);

    /// <summary>Cells whose text does not contain <paramref name="text"/>.</summary>
    /// <exception cref="ArgumentException">The text is empty.</exception>
    public static ConditionalFormat NotContainsText(string text, ConditionalStyle style) => Textual(ConditionalFormatType.NotContainsText, text, style);

    /// <summary>Cells whose text begins with <paramref name="text"/>, ignoring case.</summary>
    /// <exception cref="ArgumentException">The text is empty.</exception>
    public static ConditionalFormat BeginsWith(string text, ConditionalStyle style) => Textual(ConditionalFormatType.BeginsWith, text, style);

    /// <summary>Cells whose text ends with <paramref name="text"/>, ignoring case.</summary>
    /// <exception cref="ArgumentException">The text is empty.</exception>
    public static ConditionalFormat EndsWith(string text, ConditionalStyle style) => Textual(ConditionalFormatType.EndsWith, text, style);

    /// <summary>Empty cells and cells holding only spaces.</summary>
    public static ConditionalFormat Blanks(ConditionalStyle style) => Plain(ConditionalFormatType.Blanks, style);

    /// <summary>Cells that are not blank.</summary>
    public static ConditionalFormat NoBlanks(ConditionalStyle style) => Plain(ConditionalFormatType.NoBlanks, style);

    /// <summary>Cells holding an error.</summary>
    public static ConditionalFormat Errors(ConditionalStyle style) => Plain(ConditionalFormatType.Errors, style);

    /// <summary>Cells not holding an error.</summary>
    public static ConditionalFormat NoErrors(ConditionalStyle style) => Plain(ConditionalFormatType.NoErrors, style);

    /// <summary>Dates in a period around today, as Excel's A Date Occurring.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The period is not a defined value.</exception>
    public static ConditionalFormat DateOccurring(TimePeriod period, ConditionalStyle style)
    {
        if (period is < TimePeriod.Yesterday or > TimePeriod.NextMonth)
            throw new ArgumentOutOfRangeException(nameof(period), period, "Not a time period.");
        return new ConditionalFormat(ConditionalFormatType.TimePeriod, Checked(style)) { Period = period };
    }

    private static ConditionalFormat Plain(ConditionalFormatType type, ConditionalStyle style) => new(type, Checked(style));

    private static ConditionalStyle Checked(ConditionalStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);
        return style;
    }

    private static ConditionalFormat Ranked(ConditionalFormatType type, int rank, ConditionalStyle style, bool percent)
    {
        if (rank < 1 || rank > (percent ? 100 : 1000))
            throw new ArgumentOutOfRangeException(nameof(rank), rank, percent ? "A percentage goes from 1 to 100." : "A rank goes from 1 to 1000.");
        return new ConditionalFormat(type, Checked(style)) { Rank = rank, Percent = percent };
    }

    private static ConditionalFormat Average(ConditionalFormatType type, ConditionalStyle style, bool orEqual, int standardDeviations)
    {
        if (standardDeviations is < 0 or > 3)
            throw new ArgumentOutOfRangeException(nameof(standardDeviations), standardDeviations, "Excel takes 0 to 3 standard deviations.");
        return new ConditionalFormat(type, Checked(style)) { OrEqual = orEqual, StandardDeviations = standardDeviations };
    }

    private static ConditionalFormat Textual(ConditionalFormatType type, string text, ConditionalStyle style)
    {
        ArgumentException.ThrowIfNullOrEmpty(text);
        return new ConditionalFormat(type, Checked(style)) { Text = text };
    }

    /// <summary>A rule as a file has it, with every setting given.</summary>
    internal static ConditionalFormat Read(ConditionalFormatType type, ConditionalStyle? style, ConditionalOperator comparison, string? formula1,
        string? formula2, string? text, int rank, bool percent, bool orEqual, int standardDeviations, TimePeriod period, bool stopIfTrue) =>
        new(type, style)
        {
            Operator = comparison,
            Formula1 = formula1,
            Formula2 = formula2,
            Text = text,
            Rank = rank,
            Percent = percent,
            OrEqual = orEqual,
            StandardDeviations = standardDeviations,
            Period = period,
            StopIfTrue = stopIfTrue,
        };

    /// <summary>The rule with its formulas written for another first cell.</summary>
    internal ConditionalFormat WithFormulas(string? formula1, string? formula2) =>
        HasFormulas
            ? Read(Type, Style, Operator, formula1, formula2, Text, Rank, Percent, OrEqual, StandardDeviations, Period, StopIfTrue)
            : this;

    /// <summary>Whether the rule takes formulas of its own: those whose text it shows.</summary>
    internal bool HasFormulas => Type is ConditionalFormatType.CellValue or ConditionalFormatType.Expression;

    /// <summary>Whether the other rule has the same settings and style; texts compared ordinally.</summary>
    public bool Equals(ConditionalFormat? other) =>
        other is not null && Type == other.Type && Operator == other.Operator
        && string.Equals(Formula1, other.Formula1, StringComparison.Ordinal) && string.Equals(Formula2, other.Formula2, StringComparison.Ordinal)
        && string.Equals(Text, other.Text, StringComparison.Ordinal) && Rank == other.Rank && Percent == other.Percent
        && OrEqual == other.OrEqual && StandardDeviations == other.StandardDeviations && Period == other.Period
        && Equals(Style, other.Style) && StopIfTrue == other.StopIfTrue;

    /// <summary>A hash code consistent with <see cref="Equals(ConditionalFormat)"/>.</summary>
    public override int GetHashCode() => HashCode.Combine(Type, Operator, Formula1, Formula2, Text, Rank, Style, StopIfTrue);

    /// <summary>The type, and the formula or text when the rule has one.</summary>
    public override string ToString() => Formula1 is not null ? $"{Type} {Formula1}" : Text is not null ? $"{Type} \"{Text}\"" : Type.ToString();
}
