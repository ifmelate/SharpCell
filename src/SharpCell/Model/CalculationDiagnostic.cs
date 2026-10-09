namespace SharpCell;

public enum DiagnosticKind
{
    /// <summary>Cells reference each other in a loop; every cell in the loop got 0.</summary>
    CircularReference,

    /// <summary>A function failed unexpectedly; the cell got <c>#VALUE!</c>.</summary>
    FunctionFailure,

    /// <summary>
    /// A formula uses something SharpCell cannot evaluate, such as a table reference or (in a file)
    /// a link to another workbook; the cell got <c>#NAME?</c>.
    /// </summary>
    UnsupportedFormula,
}

/// <summary>Something calculation noticed that is not visible in cell values alone.</summary>
public sealed class CalculationDiagnostic(DiagnosticKind kind, Worksheet? sheet, string? address, string message)
{
    public DiagnosticKind Kind { get; } = kind;

    /// <summary>Sheet of the affected cell; null for <see cref="Workbook.Evaluate"/>.</summary>
    public Worksheet? Sheet { get; } = sheet;

    /// <summary>A1 address of the affected cell; null for <see cref="Workbook.Evaluate"/>.</summary>
    public string? Address { get; } = address;

    public string Message { get; } = message;

    public override string ToString() => Sheet is null ? $"{Kind}: {Message}" : $"{Kind} at {Sheet.Name}!{Address}: {Message}";
}
