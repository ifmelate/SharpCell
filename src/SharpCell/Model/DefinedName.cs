namespace SharpCell;

/// <summary>A defined name as <see cref="Workbook.DefinedNames"/> lists it.</summary>
public sealed class DefinedName
{
    internal DefinedName(string name, string formula, Worksheet? scope)
    {
        Name = name;
        Formula = formula;
        Scope = scope;
    }

    /// <summary>The name as it was defined; formulas use it ignoring case.</summary>
    public string Name { get; }

    /// <summary>The formula the name stands for, starting with <c>=</c>, such as <c>=Sheet1!$A$1:$A$10</c>.</summary>
    public string Formula { get; }

    /// <summary>The sheet the name is visible on, or null for a name of the whole workbook.</summary>
    public Worksheet? Scope { get; }

    /// <summary>The name.</summary>
    public override string ToString() => Name;
}
