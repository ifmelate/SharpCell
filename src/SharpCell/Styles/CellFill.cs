using System;
using System.Collections.Generic;
using System.Linq;

namespace SharpCell;

/// <summary>
/// The pattern of a cell's fill, as Excel's Format Cells &gt; Fill &gt; Pattern Style lists them: drawn
/// in <see cref="CellStyle.PatternColor"/> over <see cref="CellStyle.Fill"/>.
/// </summary>
public enum CellFillPattern
{
    /// <summary>No pattern: the cell is filled with <see cref="CellStyle.Fill"/>.</summary>
    Solid,

    /// <summary>75% gray dots.</summary>
    DarkGray,

    /// <summary>50% gray dots.</summary>
    MediumGray,

    /// <summary>25% gray dots.</summary>
    LightGray,

    /// <summary>12.5% gray dots.</summary>
    Gray125,

    /// <summary>6.25% gray dots.</summary>
    Gray0625,

    /// <summary>Thick horizontal stripes.</summary>
    DarkHorizontal,

    /// <summary>Thick vertical stripes.</summary>
    DarkVertical,

    /// <summary>Thick stripes down to the right.</summary>
    DarkDown,

    /// <summary>Thick stripes up to the right.</summary>
    DarkUp,

    /// <summary>A thick grid.</summary>
    DarkGrid,

    /// <summary>A thick diagonal crosshatch.</summary>
    DarkTrellis,

    /// <summary>Thin horizontal stripes.</summary>
    LightHorizontal,

    /// <summary>Thin vertical stripes.</summary>
    LightVertical,

    /// <summary>Thin stripes down to the right.</summary>
    LightDown,

    /// <summary>Thin stripes up to the right.</summary>
    LightUp,

    /// <summary>A thin grid.</summary>
    LightGrid,

    /// <summary>A thin diagonal crosshatch.</summary>
    LightTrellis,
}

/// <summary>How a gradient runs across a cell.</summary>
public enum CellGradientType
{
    /// <summary>In a straight line at <see cref="CellGradient.Angle"/>.</summary>
    Linear,

    /// <summary>Outwards from a rectangle inside the cell, given by <see cref="CellGradient.Left"/>, <see cref="CellGradient.Top"/>, <see cref="CellGradient.Right"/> and <see cref="CellGradient.Bottom"/>.</summary>
    Path,
}

/// <summary>A colour of a gradient at a position across it, from 0 (the start) to 1 (the end).</summary>
public readonly record struct GradientStop
{
    /// <summary>A stop of a gradient.</summary>
    /// <param name="position">Where the colour is, from 0 to 1.</param>
    /// <param name="color">The colour.</param>
    /// <exception cref="ArgumentOutOfRangeException">The position is outside 0 to 1.</exception>
    public GradientStop(double position, CellColor color)
    {
        Position = position is >= 0 and <= 1 ? position : throw new ArgumentOutOfRangeException(nameof(position), position, "A stop is from 0 to 1.");
        Color = color;
    }

    /// <summary>Where the colour is, from 0 to 1.</summary>
    public double Position { get; }

    /// <summary>The colour.</summary>
    public CellColor Color { get; }
}

/// <summary>
/// A gradient fill, as Excel's Format Cells &gt; Fill &gt; Fill Effects makes: two or more colours
/// blending in a line or from a rectangle outwards. A record: gradients compare by value, stops included.
/// </summary>
public sealed record CellGradient
{
    private readonly GradientStop[] _stops;
    private readonly double _left, _right, _top, _bottom;

    /// <summary>A gradient through the stops, linear at 0 degrees (left to right) unless set otherwise.</summary>
    /// <param name="stops">Two or more stops; they are kept in order of position.</param>
    /// <exception cref="ArgumentException">There are fewer than two stops.</exception>
    public CellGradient(IEnumerable<GradientStop> stops)
    {
        ArgumentNullException.ThrowIfNull(stops);
        _stops = [.. stops.OrderBy(stop => stop.Position)];
        if (_stops.Length < 2)
            throw new ArgumentException("A gradient needs at least two stops.", nameof(stops));
    }

    /// <summary>A linear gradient from one colour to another, as Excel's two-colour shading styles.</summary>
    /// <param name="from">The colour at the start.</param>
    /// <param name="to">The colour at the end.</param>
    /// <param name="angle">The direction in degrees: 0 left to right, 90 top to bottom, 45 and 135 diagonally.</param>
    public static CellGradient Linear(CellColor from, CellColor to, double angle = 0) =>
        new([new GradientStop(0, from), new GradientStop(1, to)]) { Angle = angle };

    /// <summary>The stops in order of position.</summary>
    public IReadOnlyList<GradientStop> Stops => _stops;

    /// <summary>How the gradient runs; <see cref="CellGradientType.Linear"/> unless set otherwise.</summary>
    public CellGradientType Type { get; init; }

    /// <summary>The direction of a linear gradient in degrees, clockwise from left to right.</summary>
    public double Angle { get; init; }

    /// <summary>Where a path gradient's inner rectangle starts from the left, 0 to 1 of the cell's width.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is outside 0 to 1.</exception>
    public double Left { get => _left; init => _left = Check(value); }

    /// <summary>Where the inner rectangle ends from the right, 0 to 1 of the cell's width.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is outside 0 to 1.</exception>
    public double Right { get => _right; init => _right = Check(value); }

    /// <summary>Where the inner rectangle starts from the top, 0 to 1 of the cell's height.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is outside 0 to 1.</exception>
    public double Top { get => _top; init => _top = Check(value); }

    /// <summary>Where the inner rectangle ends from the bottom, 0 to 1 of the cell's height.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is outside 0 to 1.</exception>
    public double Bottom { get => _bottom; init => _bottom = Check(value); }

    /// <summary>Whether the gradients are the same, stop by stop.</summary>
    public bool Equals(CellGradient? other) =>
        other is not null && Type == other.Type && Angle.Equals(other.Angle) && _left.Equals(other._left) && _right.Equals(other._right)
        && _top.Equals(other._top) && _bottom.Equals(other._bottom) && _stops.AsSpan().SequenceEqual(other._stops);

    /// <summary>A hash of the gradient, from its stops and shape.</summary>
    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(Type);
        hash.Add(Angle);
        foreach (var stop in _stops)
            hash.Add(stop);
        return hash.ToHashCode();
    }

    private static double Check(double value) =>
        value is >= 0 and <= 1 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "A path gradient's edges are from 0 to 1.");
}
