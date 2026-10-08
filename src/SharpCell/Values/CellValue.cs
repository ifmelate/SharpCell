using System;
using System.Globalization;
using System.Text;

namespace SharpCell;

public enum CellValueKind
{
    Empty = 0,
    Missing,
    Number,
    Text,
    Boolean,
    Error,
    Array,
    Lambda,
}

/// <summary>
/// A value produced by a cell or a formula. <c>default</c> is <see cref="Empty"/>.
/// Equality is identity (same kind, same payload, ordinal text), not Excel's <c>=</c>.
/// </summary>
public readonly struct CellValue : IEquatable<CellValue>
{
    // Number, boolean (0/1) and error code share one slot; text, array and lambda share the other.
    private readonly double _scalar;
    private readonly object? _object;

    private CellValue(CellValueKind kind, double scalar, object? obj)
    {
        Kind = kind;
        _scalar = scalar;
        _object = obj;
    }

    public CellValueKind Kind { get; }

    /// <summary>An empty cell.</summary>
    public static CellValue Empty => default;

    /// <summary>An omitted function argument, as in <c>IF(A1,,B1)</c>.</summary>
    public static CellValue Missing { get; } = new(CellValueKind.Missing, 0, null);

    public static CellValue True { get; } = new(CellValueKind.Boolean, 1, null);

    public static CellValue False { get; } = new(CellValueKind.Boolean, 0, null);

    /// <summary>
    /// A number. Excel has neither NaN, infinities nor negative zero: non-finite input
    /// becomes <c>#NUM!</c> and <c>-0</c> becomes <c>0</c>.
    /// </summary>
    public static CellValue Number(double value)
    {
        if (!double.IsFinite(value))
            return Error(ErrorKind.Num);

        // Adding positive zero turns -0 into +0 and leaves every other value unchanged.
        return new CellValue(CellValueKind.Number, value + 0.0, null);
    }

    public static CellValue Text(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return new CellValue(CellValueKind.Text, 0, value);
    }

    public static CellValue Boolean(bool value) => value ? True : False;

    public static CellValue Error(ErrorKind kind) => new(CellValueKind.Error, (int)kind, null);

    /// <summary>
    /// A rectangular array. The array is taken by reference and must not be mutated afterwards.
    /// Excel has no empty arrays, so both dimensions must be at least 1.
    /// </summary>
    public static CellValue Array(CellValue[,] values)
    {
        ArgumentNullException.ThrowIfNull(values);
        if (values.GetLength(0) == 0 || values.GetLength(1) == 0)
            throw new ArgumentException("An array must have at least one row and one column.", nameof(values));

        return new CellValue(CellValueKind.Array, 0, values);
    }

    public static CellValue Lambda(LambdaValue value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return new CellValue(CellValueKind.Lambda, 0, value);
    }

    public bool IsError => Kind == CellValueKind.Error;

    public double AsNumber() => Kind == CellValueKind.Number ? _scalar : throw WrongKind(CellValueKind.Number);

    public string AsText() => Kind == CellValueKind.Text ? (string)_object! : throw WrongKind(CellValueKind.Text);

    public bool AsBoolean() => Kind == CellValueKind.Boolean ? _scalar != 0 : throw WrongKind(CellValueKind.Boolean);

    public ErrorKind AsError() => Kind == CellValueKind.Error ? (ErrorKind)(int)_scalar : throw WrongKind(CellValueKind.Error);

    public CellValue[,] AsArray() => Kind == CellValueKind.Array ? (CellValue[,])_object! : throw WrongKind(CellValueKind.Array);

    public LambdaValue AsLambda() => Kind == CellValueKind.Lambda ? (LambdaValue)_object! : throw WrongKind(CellValueKind.Lambda);

    public static implicit operator CellValue(double value) => Number(value);

    public static implicit operator CellValue(string value) => Text(value);

    public static implicit operator CellValue(bool value) => Boolean(value);

    public bool Equals(CellValue other)
    {
        if (Kind != other.Kind)
            return false;

        return Kind switch
        {
            CellValueKind.Number or CellValueKind.Boolean or CellValueKind.Error => _scalar.Equals(other._scalar),
            CellValueKind.Text => string.Equals((string)_object!, (string)other._object!, StringComparison.Ordinal),
            CellValueKind.Array => ArraysEqual((CellValue[,])_object!, (CellValue[,])other._object!),
            CellValueKind.Lambda => ReferenceEquals(_object, other._object),
            _ => true,
        };
    }

    public override bool Equals(object? obj) => obj is CellValue other && Equals(other);

    public override int GetHashCode()
    {
        switch (Kind)
        {
            case CellValueKind.Number or CellValueKind.Boolean or CellValueKind.Error:
                return HashCode.Combine(Kind, _scalar);
            case CellValueKind.Text:
                return HashCode.Combine(Kind, StringComparer.Ordinal.GetHashCode((string)_object!));
            case CellValueKind.Array:
                var array = (CellValue[,])_object!;
                var hash = new HashCode();
                hash.Add(Kind);
                hash.Add(array.GetLength(0));
                hash.Add(array.GetLength(1));
                foreach (var item in array)
                    hash.Add(item);
                return hash.ToHashCode();
            case CellValueKind.Lambda:
                return HashCode.Combine(Kind, _object);
            default:
                return Kind.GetHashCode();
        }
    }

    public static bool operator ==(CellValue left, CellValue right) => left.Equals(right);

    public static bool operator !=(CellValue left, CellValue right) => !left.Equals(right);

    /// <summary>Formats the value the way it would be written as a formula literal.</summary>
    public override string ToString()
    {
        switch (Kind)
        {
            case CellValueKind.Number:
                return _scalar.ToString("R", CultureInfo.InvariantCulture);
            case CellValueKind.Text:
                return "\"" + ((string)_object!).Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
            case CellValueKind.Boolean:
                return _scalar != 0 ? "TRUE" : "FALSE";
            case CellValueKind.Error:
                return AsError().ToText();
            case CellValueKind.Array:
                var array = (CellValue[,])_object!;
                var sb = new StringBuilder("{");
                for (var r = 0; r < array.GetLength(0); r++)
                {
                    if (r > 0)
                        sb.Append(';');
                    for (var c = 0; c < array.GetLength(1); c++)
                    {
                        if (c > 0)
                            sb.Append(',');
                        sb.Append(array[r, c].ToString());
                    }
                }

                return sb.Append('}').ToString();
            case CellValueKind.Lambda:
                return "<lambda>";
            case CellValueKind.Missing:
                return "<missing>";
            default:
                return "<empty>";
        }
    }

    private static bool ArraysEqual(CellValue[,] a, CellValue[,] b)
    {
        if (ReferenceEquals(a, b))
            return true;
        if (a.GetLength(0) != b.GetLength(0) || a.GetLength(1) != b.GetLength(1))
            return false;

        for (var r = 0; r < a.GetLength(0); r++)
        {
            for (var c = 0; c < a.GetLength(1); c++)
            {
                if (!a[r, c].Equals(b[r, c]))
                    return false;
            }
        }

        return true;
    }

    private InvalidOperationException WrongKind(CellValueKind expected) =>
        new($"Value is {Kind}, not {expected}.");
}
