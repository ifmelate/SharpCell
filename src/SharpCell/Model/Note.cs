using System;

namespace SharpCell;

/// <summary>
/// A note on a cell: Excel's Insert Note (called a comment before Excel 365), shown in a box beside
/// the cell, by default only while the pointer rests on it. Excel's threaded comments, with replies,
/// are not notes; a file's own are kept when it is saved.
/// </summary>
public sealed class Note : IEquatable<Note>
{
    private const int MaxLength = 32767;

    /// <summary>A note with <paramref name="text"/>, written by <paramref name="author"/>.</summary>
    /// <param name="text">The note's text. Excel puts the author's name and a colon on the first line of a
    /// note it makes; SharpCell writes only the text it is given.</param>
    /// <param name="author">Who wrote the note, or null for nobody. Excel shows it in the status bar
    /// while the pointer rests on the cell.</param>
    /// <exception cref="ArgumentNullException">The text is null.</exception>
    /// <exception cref="ArgumentException">The text is longer than 32,767 characters, as Excel allows.</exception>
    public Note(string text, string? author = null)
    {
        ArgumentNullException.ThrowIfNull(text);
        if (text.Length > MaxLength)
            throw new ArgumentException($"A note is at most {MaxLength} characters.", nameof(text));
        Text = text;
        Author = author;
    }

    /// <summary>The text, with <c>\n</c> between lines.</summary>
    public string Text { get; }

    /// <summary>Who wrote the note, or null for nobody.</summary>
    public string? Author { get; }

    /// <summary>
    /// Whether the note is always shown, as Excel's Show/Hide Note makes it; otherwise it is shown
    /// only while the pointer rests on its cell.
    /// </summary>
    public bool IsVisible { get; init; }

    /// <summary>Whether the other note has the same text, author and visibility, compared ordinally.</summary>
    public bool Equals(Note? other) =>
        other is not null && string.Equals(Text, other.Text, StringComparison.Ordinal)
        && string.Equals(Author, other.Author, StringComparison.Ordinal) && IsVisible == other.IsVisible;

    /// <summary>Whether <paramref name="obj"/> is the same note.</summary>
    public override bool Equals(object? obj) => Equals(obj as Note);

    /// <summary>A hash code consistent with <see cref="Equals(Note)"/>.</summary>
    public override int GetHashCode() => HashCode.Combine(Text, Author, IsVisible);

    /// <summary>The text.</summary>
    public override string ToString() => Text;
}
