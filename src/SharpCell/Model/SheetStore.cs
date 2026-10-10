using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell;

internal readonly record struct StoredCell(int Row, int Column, CellData Data);

/// <summary>
/// Sparse cell storage: a dictionary of rows plus a sorted set of row numbers, so a range such as
/// <c>A:A</c> is walked by visiting only rows that exist. Each row keeps its columns sorted.
/// </summary>
internal sealed class SheetStore
{
    private readonly Dictionary<int, RowData> _rows = [];
    private readonly SortedSet<int> _rowNumbers = [];

    public int Count { get; private set; }

    /// <summary>Changes when a cell is added or removed, so a public enumeration can detect it.</summary>
    public int Version { get; private set; }

    /// <summary>The first and last row and column holding a cell; null when the store is empty.</summary>
    public Area? Bounds
    {
        get
        {
            if (_rowNumbers.Count == 0)
                return null;
            int first = int.MaxValue, last = 0;
            foreach (var data in _rows.Values)
            {
                first = System.Math.Min(first, data.ColumnAt(0));
                last = System.Math.Max(last, data.ColumnAt(data.Count - 1));
            }

            return new Area(_rowNumbers.Min, first, _rowNumbers.Max, last);
        }
    }

    public CellData? Get(int row, int column) => _rows.TryGetValue(row, out var data) ? data.Get(column) : null;

    public CellData GetOrCreate(int row, int column)
    {
        if (!_rows.TryGetValue(row, out var data))
        {
            data = new RowData();
            _rows.Add(row, data);
            _rowNumbers.Add(row);
        }

        var cell = data.GetOrCreate(column, out var created);
        if (created)
        {
            Count++;
            Version++;
        }
        return cell;
    }

    public void Remove(int row, int column)
    {
        if (!_rows.TryGetValue(row, out var data) || !data.Remove(column))
            return;

        Count--;
        Version++;
        if (data.IsEmpty)
        {
            _rows.Remove(row);
            _rowNumbers.Remove(row);
        }
    }

    /// <summary>
    /// Fills this empty store with copies of another store's cells. Only reads the source, so
    /// several threads may copy one store at once.
    /// </summary>
    public void CopyFrom(SheetStore source, System.Func<CellData, CellData> copy)
    {
        foreach (var row in source._rowNumbers)
            _rows.Add(row, source._rows[row].Copy(copy));
        _rowNumbers.UnionWith(source._rowNumbers);
        Count = source.Count;
    }

    /// <summary>Existing cells inside the rectangle, row by row, columns ascending. Corners must be ordered.</summary>
    public IEnumerable<StoredCell> Enumerate(int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        foreach (var row in _rowNumbers.GetViewBetween(firstRow, lastRow))
        {
            var data = _rows[row];
            for (var i = data.LowerBound(firstColumn); i < data.Count && data.ColumnAt(i) <= lastColumn; i++)
                yield return new StoredCell(row, data.ColumnAt(i), data.CellAt(i));
        }
    }

    private sealed class RowData
    {
        private readonly List<int> _columns = [];
        private readonly List<CellData> _cells = [];

        public int Count => _columns.Count;

        public bool IsEmpty => _columns.Count == 0;

        public int ColumnAt(int index) => _columns[index];

        public CellData CellAt(int index) => _cells[index];

        public RowData Copy(System.Func<CellData, CellData> copy)
        {
            var row = new RowData();
            row._columns.AddRange(_columns);
            foreach (var cell in _cells)
                row._cells.Add(copy(cell));
            return row;
        }

        public CellData? Get(int column)
        {
            var index = _columns.BinarySearch(column);
            return index >= 0 ? _cells[index] : null;
        }

        public CellData GetOrCreate(int column, out bool created)
        {
            var index = _columns.BinarySearch(column);
            if (index >= 0)
            {
                created = false;
                return _cells[index];
            }

            var cell = new CellData();
            _columns.Insert(~index, column);
            _cells.Insert(~index, cell);
            created = true;
            return cell;
        }

        public bool Remove(int column)
        {
            var index = _columns.BinarySearch(column);
            if (index < 0)
                return false;

            _columns.RemoveAt(index);
            _cells.RemoveAt(index);
            return true;
        }

        public int LowerBound(int column)
        {
            var index = _columns.BinarySearch(column);
            return index >= 0 ? index : ~index;
        }
    }
}
