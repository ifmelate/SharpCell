using System;
using System.Globalization;

namespace SharpCell;

/// <summary>
/// Iterative calculation, Excel's "Enable iterative calculation": circular references are
/// calculated over and over until they settle, instead of giving 0.
/// </summary>
public sealed class IterationSettings : IEquatable<IterationSettings>
{
    /// <summary>Settings for iterative calculation; the defaults are Excel's, with iteration off.</summary>
    /// <param name="Enabled">Whether circular references are calculated iteratively.</param>
    /// <param name="MaxIterations">The most passes over a circular reference, 1 to 32,767.</param>
    /// <param name="MaxChange">A circular reference has settled when no value of it changed by more than this in a pass; 0 or more.</param>
    /// <exception cref="ArgumentOutOfRangeException">A limit is outside what Excel allows.</exception>
    public IterationSettings(bool Enabled = false, int MaxIterations = 100, double MaxChange = 0.001)
    {
        if (MaxIterations is < 1 or > 32767)
            throw new ArgumentOutOfRangeException(nameof(MaxIterations), MaxIterations, "Excel allows 1 to 32,767 iterations.");
        if (!(MaxChange >= 0) || double.IsInfinity(MaxChange))
            throw new ArgumentOutOfRangeException(nameof(MaxChange), MaxChange, "The maximum change must be a finite number, 0 or more.");
        this.Enabled = Enabled;
        this.MaxIterations = MaxIterations;
        this.MaxChange = MaxChange;
    }

    /// <summary>Whether circular references are calculated iteratively.</summary>
    public bool Enabled { get; }

    /// <summary>The most passes over a circular reference; 100 by default.</summary>
    public int MaxIterations { get; }

    /// <summary>A circular reference has settled when no value changed by more than this in a pass; 0.001 by default.</summary>
    public double MaxChange { get; }

    /// <summary>Whether the other settings are the same.</summary>
    public bool Equals(IterationSettings? other) =>
        other is not null && other.Enabled == Enabled && other.MaxIterations == MaxIterations && other.MaxChange.Equals(MaxChange);

    /// <summary>Whether <paramref name="obj"/> is the same settings.</summary>
    public override bool Equals(object? obj) => Equals(obj as IterationSettings);

    /// <summary>A hash code consistent with <see cref="Equals(IterationSettings)"/>.</summary>
    public override int GetHashCode() => HashCode.Combine(Enabled, MaxIterations, MaxChange);

    /// <summary>The settings as text, such as <c>Enabled, 100 iterations, change 0.001</c>.</summary>
    public override string ToString() =>
        $"{(Enabled ? "Enabled" : "Disabled")}, {MaxIterations} iterations, change {MaxChange.ToString("R", CultureInfo.InvariantCulture)}";
}
