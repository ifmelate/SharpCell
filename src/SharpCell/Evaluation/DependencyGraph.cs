using System;
using System.Collections.Generic;

namespace SharpCell.Evaluation;

/// <summary>
/// Who reads what. Single cells are a dictionary from cell to readers; ranges go to a per-sheet
/// <see cref="RangeIndex"/>; names map to the cells that used them. Registrations come from what
/// a formula actually read in its last evaluation, so an untaken IF branch is not a dependency.
/// </summary>
internal sealed class DependencyGraph
{
    private readonly Dictionary<CellKey, HashSet<CellKey>> _cells = [];
    private readonly Dictionary<Worksheet, RangeIndex> _ranges = [];
    private readonly Dictionary<string, HashSet<CellKey>> _names = new(StringComparer.Ordinal);

    public int RangeCount
    {
        get
        {
            var count = 0;
            foreach (var index in _ranges.Values)
                count += index.Count;
            return count;
        }
    }

    public void Register(CellKey reader, Dependencies dependencies)
    {
        foreach (var (sheet, area) in dependencies.Areas)
        {
            if (area.IsSingleCell)
            {
                var key = new CellKey(sheet, area.FirstRow, area.FirstColumn);
                if (!_cells.TryGetValue(key, out var readers))
                    _cells[key] = readers = [];
                readers.Add(reader);
            }
            else
            {
                if (!_ranges.TryGetValue(sheet, out var index))
                    _ranges[sheet] = index = new RangeIndex();
                index.Add(area, reader);
            }
        }

        foreach (var name in dependencies.Names)
        {
            if (!_names.TryGetValue(name, out var users))
                _names[name] = users = [];
            users.Add(reader);
        }
    }

    public void Unregister(CellKey reader, Dependencies dependencies)
    {
        foreach (var (sheet, area) in dependencies.Areas)
        {
            if (area.IsSingleCell)
            {
                var key = new CellKey(sheet, area.FirstRow, area.FirstColumn);
                if (_cells.TryGetValue(key, out var readers) && readers.Remove(reader) && readers.Count == 0)
                    _cells.Remove(key);
            }
            else if (_ranges.TryGetValue(sheet, out var index))
            {
                index.Remove(area, reader);
            }
        }

        foreach (var name in dependencies.Names)
        {
            if (_names.TryGetValue(name, out var users) && users.Remove(reader) && users.Count == 0)
                _names.Remove(name);
        }
    }

    public void ForEachReader(CellKey cell, Action<CellKey> visit)
    {
        if (_cells.TryGetValue(cell, out var readers))
        {
            foreach (var reader in readers)
                visit(reader);
        }

        if (_ranges.TryGetValue(cell.Sheet, out var index))
            index.Query(cell.Row, cell.Column, visit);
    }

    public CellKey[] UsersOf(string upperName) => _names.TryGetValue(upperName, out var users) ? [.. users] : [];
}
