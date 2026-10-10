using System.Collections.Generic;
using SharpCell.Parsing;

namespace SharpCell;

/// <summary>A defined name. The formula is parsed with origin A1, as names are stored in files.</summary>
internal sealed class NameDefinition(string name, string spelling, string text, FormulaNode formula, Worksheet? scope)
{
    /// <summary>The name in upper case, as names are compared.</summary>
    public string Name { get; } = name;

    /// <summary>The name as it was defined.</summary>
    public string Spelling { get; } = spelling;

    public string Text { get; } = text;

    public FormulaNode Formula { get; } = formula;

    /// <summary>The sheet the name is local to, or null for a workbook-level name.</summary>
    public Worksheet? Scope { get; } = scope;
}

/// <summary>
/// Defined names keyed by scope and upper-case name; names are case-insensitive. Names keep the
/// order they were first defined in; redefining one keeps its place.
/// </summary>
internal sealed class NameTable
{
    private readonly Dictionary<(Worksheet? Scope, string Name), NameDefinition> _names = [];
    private readonly List<(Worksheet? Scope, string Name)> _order = [];

    public IEnumerable<NameDefinition> All
    {
        get
        {
            foreach (var key in _order)
                yield return _names[key];
        }
    }

    public void Set(NameDefinition definition)
    {
        var key = (definition.Scope, definition.Name);
        if (!_names.ContainsKey(key))
            _order.Add(key);
        _names[key] = definition;
    }

    public bool ContainsInAnyScope(string upperName)
    {
        foreach (var key in _names.Keys)
        {
            if (key.Name == upperName)
                return true;
        }

        return false;
    }

    public bool TryGet(string upperName, Worksheet? scope, out NameDefinition? definition) =>
        _names.TryGetValue((scope, upperName), out definition);
}
