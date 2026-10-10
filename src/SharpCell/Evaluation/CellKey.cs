using System;
using System.Collections.Generic;

namespace SharpCell.Evaluation;

internal readonly record struct CellKey(Worksheet Sheet, int Row, int Column)
{
    public CellAddress Address => new(Row, Column);

    public CellData? Data => Sheet.Store.Get(Row, Column);

    public override string ToString() => $"{Sheet.Name}!{Address}";
}

/// <summary>What one evaluation of a formula read: areas (cells and ranges) and defined names.</summary>
internal sealed class Dependencies
{
    public HashSet<SheetArea> Areas { get; } = [];

    /// <summary>Upper-case names, including names that were not defined at the time.</summary>
    public HashSet<string> Names { get; } = new(StringComparer.Ordinal);

    public Dependencies CopyFor(Func<Worksheet, Worksheet> map)
    {
        var copy = new Dependencies();
        foreach (var (sheet, area) in Areas)
            copy.Areas.Add(new SheetArea(map(sheet), area));
        copy.Names.UnionWith(Names);
        return copy;
    }
}
