using System;
using System.Collections.Generic;

namespace SharpCell.Evaluation;

/// <summary>
/// Which formulas read a range that contains a given cell, for one sheet. A static interval tree
/// over row intervals (an implicit balanced tree on entries sorted by first row, each node
/// knowing the largest last row below it); <c>A:A</c> is one entry, not a million cells.
/// Changes since the last build sit in small side sets and the tree is rebuilt once they grow.
/// </summary>
internal sealed class RangeIndex
{
    private readonly HashSet<(Area Area, CellKey Cell)> _entries = [];
    private readonly HashSet<(Area Area, CellKey Cell)> _added = [];
    private readonly HashSet<(Area Area, CellKey Cell)> _removed = [];
    private (Area Area, CellKey Cell)[] _sorted = [];
    private int[] _maxLastRow = [];

    public int Count => _entries.Count;

    public void Add(Area area, CellKey cell)
    {
        var entry = (area, cell);
        if (!_entries.Add(entry))
            return;
        if (!_removed.Remove(entry))
            _added.Add(entry);
    }

    public void Remove(Area area, CellKey cell)
    {
        var entry = (area, cell);
        if (!_entries.Remove(entry))
            return;
        if (!_added.Remove(entry))
            _removed.Add(entry);
    }

    public void Query(int row, int column, Action<CellKey> visit)
    {
        if (_added.Count + _removed.Count > 32 + _sorted.Length / 4)
            Rebuild();

        Query(0, _sorted.Length, row, column, visit);
        foreach (var (area, cell) in _added)
        {
            if (area.Contains(row, column))
                visit(cell);
        }
    }

    private void Query(int low, int high, int row, int column, Action<CellKey> visit)
    {
        // Recurses into the left half; the right half is a loop. Depth is the tree height, log n.
        while (low < high)
        {
            var mid = (low + high) >>> 1;
            if (_maxLastRow[mid] < row)
                return;

            Query(low, mid, row, column, visit);
            var entry = _sorted[mid];
            if (entry.Area.FirstRow > row)
                return;
            if (entry.Area.Contains(row, column) && !_removed.Contains(entry))
                visit(entry.Cell);
            low = mid + 1;
        }
    }

    private void Rebuild()
    {
        _sorted = [.. _entries];
        Array.Sort(_sorted, (a, b) => a.Area.FirstRow.CompareTo(b.Area.FirstRow));
        _maxLastRow = new int[_sorted.Length];
        Build(0, _sorted.Length);
        _added.Clear();
        _removed.Clear();
    }

    private int Build(int low, int high)
    {
        if (low >= high)
            return int.MinValue;

        var mid = (low + high) >>> 1;
        var max = Math.Max(_sorted[mid].Area.LastRow, Math.Max(Build(low, mid), Build(mid + 1, high)));
        _maxLastRow[mid] = max;
        return max;
    }
}
