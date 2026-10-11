using System;

namespace SharpCell;

/// <summary>
/// A hyperlink on a cell: a web or e-mail address, a file, or a place in the workbook. The cell's
/// value is what Excel shows; the link neither changes it nor applies the Hyperlink cell style.
/// </summary>
public sealed class Hyperlink : IEquatable<Hyperlink>
{
    private const int MaxTooltipLength = 255;

    /// <summary>A link to <paramref name="target"/>, with an optional tooltip.</summary>
    /// <param name="target">
    /// A web or e-mail address or a file (<c>https://example.com</c>, <c>mailto:a@example.com</c>, <c>report.pdf</c>),
    /// or a place in the workbook written as Excel's HYPERLINK function takes it: <c>#Sheet2!A1</c>,
    /// <c>#'Q1 Sales'!A1</c> or <c>#MyName</c>.
    /// </param>
    /// <param name="tooltip">Excel's ScreenTip, shown when the pointer rests on the cell; at most 255 characters, or null for none.</param>
    /// <exception cref="ArgumentNullException">The target is null.</exception>
    /// <exception cref="ArgumentException">The target is empty, or the tooltip is longer than 255 characters.</exception>
    public Hyperlink(string target, string? tooltip = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(target);
        if (tooltip is { Length: > MaxTooltipLength })
            throw new ArgumentException($"A tooltip is at most {MaxTooltipLength} characters.", nameof(tooltip));
        Target = target;
        Tooltip = tooltip;
    }

    /// <summary>Where the link goes, as given to the constructor.</summary>
    public string Target { get; }

    /// <summary>The ScreenTip, or null for none.</summary>
    public string? Tooltip { get; }

    /// <summary>Whether the link goes to a place in this workbook, that is, the target starts with <c>#</c>.</summary>
    public bool IsInWorkbook => Target[0] == '#';

    /// <summary>Whether the other link has the same target and tooltip, compared ordinally.</summary>
    public bool Equals(Hyperlink? other) =>
        other is not null && string.Equals(Target, other.Target, StringComparison.Ordinal) && string.Equals(Tooltip, other.Tooltip, StringComparison.Ordinal);

    /// <summary>Whether <paramref name="obj"/> is the same link.</summary>
    public override bool Equals(object? obj) => Equals(obj as Hyperlink);

    /// <summary>A hash code consistent with <see cref="Equals(Hyperlink)"/>.</summary>
    public override int GetHashCode() => HashCode.Combine(Target, Tooltip);

    /// <summary>The target.</summary>
    public override string ToString() => Target;
}
