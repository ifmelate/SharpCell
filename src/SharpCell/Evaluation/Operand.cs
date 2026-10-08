namespace SharpCell.Evaluation;

/// <summary>The result of evaluating a node: a value, or a reference that has not been read yet.</summary>
internal readonly struct Operand
{
    private Operand(CellValue value, Reference? reference)
    {
        Value = value;
        Reference = reference;
    }

    /// <summary>The value; meaningful only when <see cref="Reference"/> is null.</summary>
    public CellValue Value { get; }

    public Reference? Reference { get; }

    public bool IsReference => Reference is not null;

    public static Operand Of(Reference reference) => new(default, reference);

    public static implicit operator Operand(CellValue value) => new(value, null);
}
