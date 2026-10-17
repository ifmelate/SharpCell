using System;
using System.Collections.Generic;

namespace SharpCell.Parsing;

/// <summary>
/// Rewrites the sheet names in a formula's references, as Excel does when a sheet is renamed or
/// removed. Text inside INDIRECT and links to other workbooks do not change.
/// </summary>
internal static class SheetReferences
{
    /// <summary>The formula with every reference naming <paramref name="from"/> naming <paramref name="to"/>; the same instance when none does.</summary>
    public static FormulaNode Rename(FormulaNode node, string from, string to) =>
        new Mapper(prefix =>
        {
            var first = Is(prefix.First, from);
            var last = prefix.Last is not null && Is(prefix.Last, from);
            return first || last ? new SheetPrefix(first ? to : prefix.First, last ? to : prefix.Last) : prefix;
        }).Visit(node);

    /// <summary>
    /// The formula with references into a removed sheet broken, as Excel breaks them: <c>Sheet2!A1</c>
    /// becomes <c>#REF!A1</c>. A reference through several sheets whose first or last sheet is
    /// removed keeps the sheets left between them: <c>Sheet1:Sheet3!A1</c> without Sheet1 is
    /// <c>Sheet2:Sheet3!A1</c>.
    /// </summary>
    /// <param name="node">The formula.</param>
    /// <param name="removed">The removed sheet's name.</param>
    /// <param name="order">The sheets' names in order, the removed one still among them.</param>
    public static FormulaNode Remove(FormulaNode node, string removed, IReadOnlyList<string> order) =>
        new Mapper(prefix => Removed(prefix, removed, order)).Visit(node);

    /// <summary>Applies <paramref name="map"/> to every sheet prefix; a prefix it returns as is stays.</summary>
    public static FormulaNode Map(FormulaNode node, Func<SheetPrefix, SheetPrefix> map) => new Mapper(map).Visit(node);

    /// <summary>The prefix once <paramref name="removed"/> is gone from the sheets in <paramref name="order"/>.</summary>
    public static SheetPrefix Removed(SheetPrefix prefix, string removed, IReadOnlyList<string> order)
    {
        if (prefix.Last is null)
            return Is(prefix.First, removed) ? SheetPrefix.Deleted : prefix;
        bool first = Is(prefix.First, removed), last = Is(prefix.Last, removed);
        if (!first && !last)
            return prefix;
        int i = IndexOf(order, prefix.First), j = IndexOf(order, prefix.Last);
        if (i < 0 || j < 0 || i == j)
            return SheetPrefix.Deleted;
        // The removed end moves one sheet towards the other end.
        var step = j > i ? 1 : -1;
        return first ? prefix with { First = order[i + step] } : prefix with { Last = order[j - step] };
    }

    private static int IndexOf(IReadOnlyList<string> order, string name)
    {
        for (var i = 0; i < order.Count; i++)
        {
            if (Is(order[i], name))
                return i;
        }

        return -1;
    }

    private static bool Is(string name, string sheet) => string.Equals(name, sheet, StringComparison.OrdinalIgnoreCase);

    private sealed class Mapper(Func<SheetPrefix, SheetPrefix> map) : FormulaRewriter
    {
        protected override FormulaNode? Replace(FormulaNode node)
        {
            switch (node)
            {
                case ReferenceNode { Sheet: { } sheet } reference:
                    return map(sheet) is var mapped && mapped != sheet ? new ReferenceNode(mapped, reference.Area) : reference;
                case RefErrorNode error:
                    return map(error.Sheet) is var broken && broken != error.Sheet ? new RefErrorNode(broken) : error;
                case NameNode { Sheet: { } sheet } name:
                    var scope = map(sheet);
                    if (scope == sheet)
                        return name;
                    // A sheet's own name is gone with the sheet.
                    return scope.IsDeleted ? new ErrorNode(ErrorKind.Ref) : new NameNode(scope, name.Name);
                default:
                    return null;
            }
        }
    }
}
