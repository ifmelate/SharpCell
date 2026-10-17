using System;
using System.Collections.Generic;
using System.Linq;

namespace SharpCell;

/// <summary>How a column of an <see cref="AutoFilter"/> picks the rows it shows, as Excel's filter menu offers it.</summary>
public enum ColumnFilterType
{
    /// <summary>Rows whose cell shows one of <see cref="ColumnFilter.Values"/> or <see cref="ColumnFilter.Dates"/>, or is blank when <see cref="ColumnFilter.Blanks"/>: the menu's check boxes.</summary>
    Values,

    /// <summary>Rows whose cell meets one or two conditions: Text Filters and Number Filters &gt; Custom Filter.</summary>
    Custom,

    /// <summary>Rows with the highest numbers: Number Filters &gt; Top 10.</summary>
    Top,

    /// <summary>Rows with the lowest numbers.</summary>
    Bottom,

    /// <summary>Rows above or below the average, or with dates in a period around today: <see cref="ColumnFilter.Dynamic"/>.</summary>
    Dynamic,

    /// <summary>
    /// Filter by Color (a fill, a font colour or an icon), read from a file. SharpCell keeps it when
    /// saving but cannot apply it, since cell colours from conditional formats are not worked out.
    /// </summary>
    ByColor,
}

/// <summary>How a <see cref="FilterCondition"/> compares a cell with its value.</summary>
public enum FilterOperator
{
    /// <summary>The cell shows the value, ignoring case; <c>*</c> and <c>?</c> are wildcards, <c>~</c> escapes them.</summary>
    Equal,

    /// <summary>The cell does not show the value.</summary>
    NotEqual,

    /// <summary>Greater than the value: a number than a number, a text than a text.</summary>
    GreaterThan,

    /// <summary>Greater than or equal to the value.</summary>
    GreaterThanOrEqual,

    /// <summary>Less than the value.</summary>
    LessThan,

    /// <summary>Less than or equal to the value.</summary>
    LessThanOrEqual,
}

/// <summary>One condition of a <see cref="ColumnFilterType.Custom"/> filter.</summary>
/// <param name="Operator">How the cell is compared.</param>
/// <param name="Value">A number written with <c>.</c> for decimals (dates as serial numbers), or a text.</param>
public readonly record struct FilterCondition(FilterOperator Operator, string Value);

/// <summary>The rows a <see cref="ColumnFilterType.Dynamic"/> filter shows: around the average, or dates around today.</summary>
public enum DynamicFilter
{
    /// <summary>Numbers above the column's average.</summary>
    AboveAverage,

    /// <summary>Numbers below the column's average.</summary>
    BelowAverage,

    /// <summary>Dates yesterday.</summary>
    Yesterday,

    /// <summary>Dates today.</summary>
    Today,

    /// <summary>Dates tomorrow.</summary>
    Tomorrow,

    /// <summary>Dates in the week before this one, Sunday to Saturday.</summary>
    LastWeek,

    /// <summary>Dates in this week, Sunday to Saturday.</summary>
    ThisWeek,

    /// <summary>Dates in the week after this one.</summary>
    NextWeek,

    /// <summary>Dates in the month before this one.</summary>
    LastMonth,

    /// <summary>Dates in this month.</summary>
    ThisMonth,

    /// <summary>Dates in the month after this one.</summary>
    NextMonth,

    /// <summary>Dates in the quarter before this one.</summary>
    LastQuarter,

    /// <summary>Dates in this quarter.</summary>
    ThisQuarter,

    /// <summary>Dates in the quarter after this one.</summary>
    NextQuarter,

    /// <summary>Dates in the year before this one.</summary>
    LastYear,

    /// <summary>Dates in this year.</summary>
    ThisYear,

    /// <summary>Dates in the year after this one.</summary>
    NextYear,

    /// <summary>Dates from the first of January to today.</summary>
    YearToDate,

    /// <summary>Dates in January to March of any year.</summary>
    Quarter1,

    /// <summary>Dates in April to June of any year.</summary>
    Quarter2,

    /// <summary>Dates in July to September of any year.</summary>
    Quarter3,

    /// <summary>Dates in October to December of any year.</summary>
    Quarter4,

    /// <summary>Dates in January of any year.</summary>
    January,

    /// <summary>Dates in February of any year.</summary>
    February,

    /// <summary>Dates in March of any year.</summary>
    March,

    /// <summary>Dates in April of any year.</summary>
    April,

    /// <summary>Dates in May of any year.</summary>
    May,

    /// <summary>Dates in June of any year.</summary>
    June,

    /// <summary>Dates in July of any year.</summary>
    July,

    /// <summary>Dates in August of any year.</summary>
    August,

    /// <summary>Dates in September of any year.</summary>
    September,

    /// <summary>Dates in October of any year.</summary>
    October,

    /// <summary>Dates in November of any year.</summary>
    November,

    /// <summary>Dates in December of any year.</summary>
    December,
}

/// <summary>How much of a <see cref="FilterDate"/> counts: a year, a month, a day, down to a second.</summary>
public enum DateGrouping
{
    /// <summary>The whole year.</summary>
    Year,

    /// <summary>The whole month.</summary>
    Month,

    /// <summary>The whole day.</summary>
    Day,

    /// <summary>The whole hour.</summary>
    Hour,

    /// <summary>The whole minute.</summary>
    Minute,

    /// <summary>The second.</summary>
    Second,
}

/// <summary>
/// A year, month, day (or hour, minute, second) a <see cref="ColumnFilterType.Values"/> filter shows,
/// as the date tree of Excel's filter menu checks it: <c>new FilterDate(DateGrouping.Month, 2024, 3)</c>
/// is March 2024. Parts finer than the grouping are left out.
/// </summary>
/// <param name="Grouping">How much of the date counts.</param>
/// <param name="Year">The year.</param>
/// <param name="Month">The month, 1 to 12.</param>
/// <param name="Day">The day of the month.</param>
/// <param name="Hour">The hour, 0 to 23.</param>
/// <param name="Minute">The minute.</param>
/// <param name="Second">The second.</param>
public readonly record struct FilterDate(DateGrouping Grouping, int Year, int Month = 1, int Day = 1, int Hour = 0, int Minute = 0, int Second = 0)
{
    /// <summary>Whether a date and time falls in this year, month, day, hour, minute or second.</summary>
    public bool Contains(DateTime value) =>
        value.Year == Year
        && (Grouping < DateGrouping.Month || value.Month == Month)
        && (Grouping < DateGrouping.Day || value.Day == Day)
        && (Grouping < DateGrouping.Hour || value.Hour == Hour)
        && (Grouping < DateGrouping.Minute || value.Minute == Minute)
        && (Grouping < DateGrouping.Second || value.Second == Second);
}

/// <summary>
/// What a column of an <see cref="AutoFilter"/> shows. Make one with a factory such as
/// <see cref="OneOf(string[])"/>, <see cref="Custom(FilterOperator, string)"/> or <see cref="Top"/>, and
/// give it to <see cref="AutoFilter.Filter"/>.
/// </summary>
public sealed record ColumnFilter
{
    private ColumnFilter(ColumnFilterType type) => Type = type;

    /// <summary>What the filter looks at.</summary>
    public ColumnFilterType Type { get; }

    /// <summary>For <see cref="ColumnFilterType.Values"/>: the texts shown, as the cells display them, ignoring case.</summary>
    public IReadOnlyList<string> Values { get; init; } = [];

    /// <summary>For <see cref="ColumnFilterType.Values"/>: the years, months or days of the dates shown.</summary>
    public IReadOnlyList<FilterDate> Dates { get; init; } = [];

    /// <summary>For <see cref="ColumnFilterType.Values"/>: whether empty cells show too, the menu's (Blanks).</summary>
    public bool Blanks { get; init; }

    /// <summary>For <see cref="ColumnFilterType.Custom"/>: the conditions, one or two.</summary>
    public IReadOnlyList<FilterCondition> Conditions { get; private init; } = [];

    /// <summary>For <see cref="ColumnFilterType.Custom"/> with two conditions: whether a row must meet both (And) or one (Or).</summary>
    public bool MatchAll { get; private init; }

    /// <summary>For <see cref="ColumnFilterType.Top"/> and <see cref="ColumnFilterType.Bottom"/>: how many rows, or what percent of them.</summary>
    public double Count { get; private init; }

    /// <summary>For <see cref="ColumnFilterType.Top"/> and <see cref="ColumnFilterType.Bottom"/>: whether <see cref="Count"/> is a percent.</summary>
    public bool Percent { get; private init; }

    /// <summary>For <see cref="ColumnFilterType.Dynamic"/>: the rows shown.</summary>
    public DynamicFilter Dynamic { get; private init; }

    /// <summary>A filter SharpCell does not apply, as a file has it, for writing it back.</summary>
    internal object? Source { get; init; }

    /// <summary>Whether the filter has criteria; a column read without any only hides its button.</summary>
    internal bool HasCriteria => Type != ColumnFilterType.Values || Values.Count > 0 || Dates.Count > 0 || Blanks;

    /// <summary>Rows whose cell shows one of the texts, as the menu's check boxes pick them.</summary>
    /// <param name="values">The texts, as the cells display them (a number with its format, such as <c>1,250.00</c>), ignoring case.</param>
    public static ColumnFilter OneOf(params string[] values)
    {
        ArgumentNullException.ThrowIfNull(values);
        if (Array.Exists(values, value => value is null))
            throw new ArgumentException("A value is null.", nameof(values));
        return new ColumnFilter(ColumnFilterType.Values) { Values = [.. values] };
    }

    /// <summary>Rows whose date falls in one of the years, months or days, as the menu's date tree picks them.</summary>
    public static ColumnFilter OneOf(params FilterDate[] dates)
    {
        ArgumentNullException.ThrowIfNull(dates);
        return new ColumnFilter(ColumnFilterType.Values) { Dates = [.. dates] };
    }

    /// <summary>Rows whose cell is empty, the menu's (Blanks) alone.</summary>
    public static ColumnFilter Blank() => new(ColumnFilterType.Values) { Blanks = true };

    /// <summary>Rows whose cell meets a condition, such as <c>Custom(FilterOperator.GreaterThan, "100")</c>.</summary>
    /// <exception cref="ArgumentException">The operator is not one of <see cref="FilterOperator"/>'s.</exception>
    public static ColumnFilter Custom(FilterOperator @operator, string value) => Custom([new FilterCondition(@operator, value)], false);

    /// <summary>Rows whose cell meets both conditions (<paramref name="matchAll"/>) or either of them.</summary>
    /// <exception cref="ArgumentException">An operator is not one of <see cref="FilterOperator"/>'s.</exception>
    public static ColumnFilter Custom(FilterCondition first, FilterCondition second, bool matchAll) => Custom([first, second], matchAll);

    private static ColumnFilter Custom(FilterCondition[] conditions, bool matchAll)
    {
        foreach (var condition in conditions)
        {
            ArgumentNullException.ThrowIfNull(condition.Value, "value");
            if (condition.Operator is < FilterOperator.Equal or > FilterOperator.LessThanOrEqual)
                throw new ArgumentException($"'{condition.Operator}' is not a filter operator.", "operator");
        }

        return new ColumnFilter(ColumnFilterType.Custom) { Conditions = conditions, MatchAll = matchAll };
    }

    /// <summary>Rows whose cell begins with a text, ignoring case: Text Filters &gt; Begins With.</summary>
    public static ColumnFilter BeginsWith(string text) => Custom(FilterOperator.Equal, Escape(text) + "*");

    /// <summary>Rows whose cell ends with a text, ignoring case.</summary>
    public static ColumnFilter EndsWith(string text) => Custom(FilterOperator.Equal, "*" + Escape(text));

    /// <summary>Rows whose cell contains a text, ignoring case.</summary>
    public static ColumnFilter Contains(string text) => Custom(FilterOperator.Equal, "*" + Escape(text) + "*");

    /// <summary>Rows whose cell does not contain a text, ignoring case.</summary>
    public static ColumnFilter DoesNotContain(string text) => Custom(FilterOperator.NotEqual, "*" + Escape(text) + "*");

    /// <summary>The rows with the highest numbers: Number Filters &gt; Top 10.</summary>
    /// <param name="count">How many rows, 1 to 500, or with <paramref name="percent"/> what percent of the rows with numbers, 1 to 100.</param>
    /// <param name="percent">Whether the count is a percent.</param>
    /// <exception cref="ArgumentOutOfRangeException">The count is outside what Excel allows.</exception>
    public static ColumnFilter Top(double count, bool percent = false) => Ranked(ColumnFilterType.Top, count, percent);

    /// <summary>The rows with the lowest numbers.</summary>
    /// <inheritdoc cref="Top"/>
    public static ColumnFilter Bottom(double count, bool percent = false) => Ranked(ColumnFilterType.Bottom, count, percent);

    private static ColumnFilter Ranked(ColumnFilterType type, double count, bool percent)
    {
        if (!(count >= 1 && count <= (percent ? 100 : 500)))
            throw new ArgumentOutOfRangeException(nameof(count), count, percent ? "A percent must be 1 to 100." : "A count must be 1 to 500.");
        return new ColumnFilter(type) { Count = count, Percent = percent };
    }

    /// <summary>The rows above or below the average, or with dates in a period around today.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not one of <see cref="DynamicFilter"/>'s.</exception>
    public static ColumnFilter Of(DynamicFilter dynamic)
    {
        if (dynamic is < DynamicFilter.AboveAverage or > DynamicFilter.December)
            throw new ArgumentOutOfRangeException(nameof(dynamic), dynamic, "Not a dynamic filter.");
        return new ColumnFilter(ColumnFilterType.Dynamic) { Dynamic = dynamic };
    }

    /// <summary>A filter by colour or icon read from a file, kept as it is.</summary>
    internal static ColumnFilter ByColor(object source) => new(ColumnFilterType.ByColor) { Source = source };

    // A text matched literally in a wildcard pattern.
    private static string Escape(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        return text.Replace("~", "~~", StringComparison.Ordinal).Replace("*", "~*", StringComparison.Ordinal).Replace("?", "~?", StringComparison.Ordinal);
    }

    /// <summary>Whether two filters pick the same rows the same way: their values and dates in the same order.</summary>
    public bool Equals(ColumnFilter? other) =>
        other is not null && Type == other.Type && Values.SequenceEqual(other.Values) && Dates.SequenceEqual(other.Dates) && Blanks == other.Blanks
        && Conditions.SequenceEqual(other.Conditions) && MatchAll == other.MatchAll && Count.Equals(other.Count) && Percent == other.Percent
        && Dynamic == other.Dynamic && Equals(Source, other.Source);

    /// <summary>A hash code agreeing with <see cref="Equals(ColumnFilter)"/>.</summary>
    public override int GetHashCode() => HashCode.Combine(Type, Values.Count, Dates.Count, Blanks, Conditions.Count, Count, Dynamic);
}
