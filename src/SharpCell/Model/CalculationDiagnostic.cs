namespace SharpCell;

/// <summary>What a <see cref="CalculationDiagnostic"/> reports.</summary>
public enum DiagnosticKind
{
    /// <summary>Cells reference each other in a loop; every cell in the loop got 0.</summary>
    CircularReference,

    /// <summary>A function failed unexpectedly; the cell got <c>#VALUE!</c>.</summary>
    FunctionFailure,

    /// <summary>
    /// A formula uses something SharpCell cannot evaluate, such as a function it does not know or
    /// (in a file) a link to another workbook; the cell got <c>#NAME?</c>.
    /// </summary>
    UnsupportedFormula,

    /// <summary>
    /// A result was larger than a limit allows, such as the cells all spills of a workbook may
    /// cover together; the cell got an error.
    /// </summary>
    LimitExceeded,
}

/// <summary>Something calculation noticed that is not visible in cell values alone.</summary>
public sealed class CalculationDiagnostic(DiagnosticKind kind, Worksheet? sheet, string? address, string message)
{
    /// <summary>What happened.</summary>
    public DiagnosticKind Kind { get; } = kind;

    /// <summary>Sheet of the affected cell; null for <see cref="Workbook.Evaluate(string)"/>.</summary>
    public Worksheet? Sheet { get; } = sheet;

    /// <summary>A1 address of the affected cell; null for <see cref="Workbook.Evaluate(string)"/>.</summary>
    public string? Address { get; } = address;

    /// <summary>A description for people, such as the cells that form a circular reference.</summary>
    public string Message { get; } = message;

    /// <summary>The kind, the cell if there is one, and the message on one line.</summary>
    public override string ToString() => Sheet is null ? $"{Kind}: {Message}" : $"{Kind} at {Sheet.Name}!{Address}: {Message}";
}
