using SharpCell.Evaluation;

namespace SharpCell.Functions;

internal enum FunctionStatus
{
    Implemented,

    /// <summary>Works, but differs from Excel in a documented way (see <see cref="FunctionInfo.Deviation"/>).</summary>
    KnownDeviation,

    /// <summary>Known to Excel but not available; calls give <c>#NAME?</c>.</summary>
    NotImplemented,
}

internal enum ArgumentKind
{
    /// <summary>
    /// A scalar. The argument is read as a value; an array (or a multi-cell range) makes the engine
    /// call the function once per element and build an array result, so functions are written for scalars.
    /// </summary>
    Value,

    /// <summary>Passed as evaluated, reference or array included: SUM, ROWS.</summary>
    Any,

    /// <summary>Evaluated only when the function asks for it: the branches of IF.</summary>
    Lazy,
}

internal delegate Operand FunctionBody(FunctionCall call);

/// <summary>A registry entry: the implementation plus the metadata the engine and the compatibility table use.</summary>
internal sealed class FunctionInfo(string name, int minArguments, int maxArguments, ArgumentKind[] kinds, FunctionBody body)
{
    public string Name { get; } = name;

    public int MinArguments { get; } = minArguments;

    public int MaxArguments { get; } = maxArguments;

    public FunctionBody Body { get; } = body;

    public bool IsVolatile { get; init; }

    public FunctionStatus Status { get; init; } = FunctionStatus.Implemented;

    public string? Deviation { get; init; }

    /// <summary>The kind of argument <paramref name="index"/>; the last kind repeats for variadic functions.</summary>
    public ArgumentKind KindAt(int index) => kinds.Length == 0 ? ArgumentKind.Any : kinds[index < kinds.Length ? index : ^1];
}
