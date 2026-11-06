using System;
using System.Collections.Generic;
using System.Linq;

namespace SharpCell;

/// <summary>A stretch of a cell's text in one font.</summary>
/// <param name="Text">The text.</param>
/// <param name="Font">Its font; null for the cell's own font.</param>
public readonly record struct TextRun(string Text, CellFont? Font = null);

/// <summary>
/// A cell's text in stretches of different fonts, as Excel keeps text formatted character by
/// character: a word in bold, a unit in superscript. The cell's value is the whole text; runs
/// without text of their own are dropped. A record: rich texts compare by value, run by run.
/// </summary>
public sealed record RichText
{
    /// <summary>The most characters a cell holds, as in Excel.</summary>
    public static int MaxLength => 32767;

    private readonly TextRun[] _runs;

    /// <summary>Rich text made of the runs, in order.</summary>
    /// <exception cref="ArgumentException">The runs hold no text, or more than 32,767 characters.</exception>
    public RichText(IEnumerable<TextRun> runs)
    {
        ArgumentNullException.ThrowIfNull(runs);
        _runs = [.. runs.Where(run => !string.IsNullOrEmpty(run.Text))];
        if (_runs.Length == 0)
            throw new ArgumentException("Rich text needs some text.", nameof(runs));
        Text = string.Concat(_runs.Select(run => run.Text));
        if (Text.Length > MaxLength)
            throw new ArgumentException($"A cell holds at most {MaxLength} characters.", nameof(runs));
    }

    /// <summary>Rich text made of the runs, in order.</summary>
    /// <exception cref="ArgumentException">The runs hold no text, or more than 32,767 characters.</exception>
    public RichText(params TextRun[] runs)
        : this((IEnumerable<TextRun>)runs)
    {
    }

    /// <summary>The runs, in order, none of them empty.</summary>
    public IReadOnlyList<TextRun> Runs => _runs;

    /// <summary>The whole text, as the cell's value has it.</summary>
    public string Text { get; }

    /// <summary>Whether the texts have the same runs.</summary>
    public bool Equals(RichText? other) => other is not null && _runs.AsSpan().SequenceEqual(other._runs);

    /// <summary>A hash of the runs.</summary>
    public override int GetHashCode()
    {
        var hash = new HashCode();
        foreach (var run in _runs)
            hash.Add(run);
        return hash.ToHashCode();
    }

    /// <summary>The whole text.</summary>
    public override string ToString() => Text;
}
