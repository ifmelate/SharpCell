using System.Collections.Generic;
using SharpCell.Parsing;

namespace SharpCell;

/// <summary>A defined name. The formula is parsed with origin A1, as names are stored in files.</summary>
internal sealed class NameDefinition(string name, string text, FormulaNode formula, Worksheet? scope)
{
    public string Name { get; } = name;

    public string Text { get; } = text;

    public FormulaNode Formula { get; } = formula;

    /// <summary>The sheet the name is local to, or null for a workbook-level name.</summary>
    public Worksheet? Scope { get; } = scope;
}

/// <summary>Defined names keyed by scope and upper-case name; names are case-insensitive.</summary>
internal sealed class NameTable
{
    private readonly Dictionary<(Worksheet? Scope, string Name), NameDefinition> _names = [];

    public void Set(NameDefinition definition) => _names[(definition.Scope, definition.Name)] = definition;

    public bool TryGet(string upperName, Worksheet? scope, out NameDefinition? definition) =>
        _names.TryGetValue((scope, upperName), out definition);
}
