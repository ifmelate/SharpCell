using System;
using System.Collections;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell;

internal readonly record struct StoredCell(int Row, int Column, CellData Data);

/// <summary>
/// Sparse cell storage. Rows are found by number through a dictionary, and kept in order in chunks
/// of at most <see cref="ChunkSize"/> rows, so a range such as <c>A:A</c> is walked by visiting only
/// rows that exist: one binary search for its first row, then row after row. Each row keeps its
/// cells sorted by column in one array.
/// </summary>
internal sealed class SheetStore
{
    // Large enough that walking rows rarely crosses into another chunk, small enough that putting a
    // row in the middle of a chunk moves little.
    private const int ChunkSize = 128;

    private readonly Dictionary<int, RowData> _rows = [];
    private Chunk[] _chunks = [];
    private int _chunkCount;

    public int Count { get; private set; }

    /// <summary>Changes when a cell is added or removed, so an enumeration can detect it.</summary>
    public int Version { get; private set; }

    /// <summary>The first and last row and column holding a cell; null when the store is empty.</summary>
    public Area? Bounds
    {
        get
        {
            if (_chunkCount == 0)
                return null;
            int first = int.MaxValue, last = 0;
            for (var c = 0; c < _chunkCount; c++)
            {
                var chunk = _chunks[c];
                for (var i = 0; i < chunk.Count; i++)
                {
                    var data = chunk.Rows[i];
                    first = Math.Min(first, data.Entries[0].Column);
                    last = Math.Max(last, data.Entries[data.Count - 1].Column);
                }
            }

            var lastChunk = _chunks[_chunkCount - 1];
            return new Area(_chunks[0].Keys[0], first, lastChunk.Keys[lastChunk.Count - 1], last);
        }
    }

    public CellData? Get(int row, int column) => _rows.TryGetValue(row, out var data) ? data.Get(column) : null;

    public CellData GetOrCreate(int row, int column)
    {
        var data = Row(row);
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
        if (data.Count == 0)
        {
            _rows.Remove(row);
            RemoveRow(row);
        }
    }

    /// <summary>
    /// Fills this empty store with copies of another store's cells. Only reads the source, so
    /// several threads may copy one store at once.
    /// </summary>
    public void CopyFrom(SheetStore source, Func<CellData, CellData> copy)
    {
        _rows.EnsureCapacity(source._rows.Count);
        _chunks = new Chunk[source._chunkCount];
        for (var c = 0; c < source._chunkCount; c++)
        {
            var from = source._chunks[c];
            var chunk = new Chunk { Count = from.Count };
            Array.Copy(from.Keys, chunk.Keys, from.Count);
            for (var i = 0; i < from.Count; i++)
            {
                var data = from.Rows[i].Copy(copy);
                chunk.Rows[i] = data;
                _rows.Add(from.Keys[i], data);
            }

            _chunks[c] = chunk;
        }

        _chunkCount = source._chunkCount;
        Count = source.Count;
        Version++;
    }

    /// <summary>Moves every cell to where <paramref name="map"/> sends it; cells it sends nowhere are dropped.</summary>
    public void Move(Func<int, int, CellAddress?> map)
    {
        var cells = new List<StoredCell>(Count);
        foreach (var cell in Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
            cells.Add(cell);
        _rows.Clear();
        _chunks = [];
        _chunkCount = 0;
        Count = 0;
        Version++;
        foreach (var (row, column, data) in cells)
        {
            if (map(row, column) is not { } to)
                continue;
            if (Row(to.Row).Put(to.Column, data))
                Count++;
        }
    }

    /// <summary>
    /// Existing cells inside the rectangle, row by row, columns ascending. Corners must be ordered.
    /// Cells may be added or removed while it runs: it goes on after the last cell it returned, so
    /// a cell is never returned twice and cells that stay are not skipped.
    /// </summary>
    public StoredCells Enumerate(int firstRow, int firstColumn, int lastRow, int lastColumn) =>
        new(this, firstRow, firstColumn, lastRow, lastColumn);

    /// <summary>The row, created (and put in order) when missing.</summary>
    private RowData Row(int row)
    {
        if (!_rows.TryGetValue(row, out var data))
        {
            // Rows written one after another tend to be as wide as the one before.
            var previous = _chunkCount == 0 ? null : _chunks[_chunkCount - 1].Rows[_chunks[_chunkCount - 1].Count - 1];
            data = new RowData();
            _rows.Add(row, data);
            InsertRow(row, data);
        }

        return data;
    }

    private void InsertRow(int row, RowData data)
    {
        if (_chunkCount == 0 || row > _chunks[_chunkCount - 1].LastKey)
        {
            // Rows mostly come in order: append, starting a new chunk when the last one is full.
            if (_chunkCount == 0 || _chunks[_chunkCount - 1].Count == ChunkSize)
                InsertChunk(_chunkCount, new Chunk());
            _chunks[_chunkCount - 1].Append(row, data);
            return;
        }

        var c = Math.Max(0, FindChunk(row));
        var chunk = _chunks[c];
        if (chunk.Count == ChunkSize)
        {
            var upper = chunk.Split();
            InsertChunk(c + 1, upper);
            if (row > upper.Keys[0])
                chunk = upper;
        }

        chunk.Insert(row, data);
    }

    private void RemoveRow(int row)
    {
        var c = FindChunk(row);
        var chunk = _chunks[c];
        chunk.Remove(row);
        if (chunk.Count > 0)
            return;
        _chunkCount--;
        Array.Copy(_chunks, c + 1, _chunks, c, _chunkCount - c);
        _chunks[_chunkCount] = null!;
    }

    private void InsertChunk(int index, Chunk chunk)
    {
        if (_chunkCount == _chunks.Length)
            Array.Resize(ref _chunks, Math.Max(4, _chunks.Length * 2));
        Array.Copy(_chunks, index, _chunks, index + 1, _chunkCount - index);
        _chunks[index] = chunk;
        _chunkCount++;
    }

    /// <summary>The index of the last chunk whose first row is at most <paramref name="row"/>; -1 when there is none.</summary>
    private int FindChunk(int row)
    {
        int low = 0, high = _chunkCount - 1;
        while (low <= high)
        {
            var middle = (low + high) >>> 1;
            if (_chunks[middle].Keys[0] <= row)
                low = middle + 1;
            else
                high = middle - 1;
        }

        return high;
    }

    /// <summary>The cells <see cref="Enumerate"/> returns; a struct, so <c>foreach</c> allocates nothing.</summary>
    internal readonly struct StoredCells(SheetStore store, int firstRow, int firstColumn, int lastRow, int lastColumn)
        : IEnumerable<StoredCell>
    {
        public Enumerator GetEnumerator() => new(store, firstRow, firstColumn, lastRow, lastColumn);

        IEnumerator<StoredCell> IEnumerable<StoredCell>.GetEnumerator() => GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    internal struct Enumerator : IEnumerator<StoredCell>
    {
        private readonly SheetStore _store;
        private readonly int _firstRow, _firstColumn, _lastRow, _lastColumn;
        private int _version;
        private bool _started;
        private Chunk? _chunk;
        private int _chunkIndex, _slot;
        private RowData? _row;
        private int _rowNumber, _cell;
        private StoredCell _current;

        internal Enumerator(SheetStore store, int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            _store = store;
            (_firstRow, _firstColumn, _lastRow, _lastColumn) = (firstRow, firstColumn, lastRow, lastColumn);
            _current = default;
        }

        public readonly StoredCell Current => _current;

        readonly object IEnumerator.Current => _current;

        public bool MoveNext()
        {
            if (!_started || _version != _store.Version)
                Seek();

            while (_row is not null)
            {
                var row = _row;
                var i = ++_cell;
                if (i < row.Count)
                {
                    ref readonly var entry = ref row.Entries[i];
                    if (entry.Column <= _lastColumn)
                    {
                        _current = new StoredCell(_rowNumber, entry.Column, entry.Data);
                        return true;
                    }
                }

                var chunk = _chunk!;
                if (++_slot == chunk.Count)
                {
                    if (++_chunkIndex == _store._chunkCount)
                    {
                        _row = null;
                        return false;
                    }

                    chunk = _chunk = _store._chunks[_chunkIndex];
                    _slot = 0;
                }

                Enter(chunk, _firstColumn);
            }

            return false;
        }

        // Finds where to go on: the first cell after the last one returned, or the first one of the range.
        private void Seek()
        {
            var store = _store;
            _version = store.Version;
            int afterRow, afterColumn;
            if (!_started)
            {
                _started = true;
                (afterRow, afterColumn) = (_firstRow, _firstColumn - 1);
            }
            else if (_row is null)
            {
                return;
            }
            else
            {
                (afterRow, afterColumn) = (_current.Row, _current.Column);
            }

            _row = null;
            var c = store.FindChunk(afterRow);
            int slot;
            if (c < 0)
            {
                c = 0;
                slot = 0;
            }
            else
            {
                slot = store._chunks[c].LowerBound(afterRow);
                if (slot == store._chunks[c].Count)
                {
                    c++;
                    slot = 0;
                }
            }

            if (c >= store._chunkCount)
                return;
            var chunk = _chunk = store._chunks[c];
            (_chunkIndex, _slot) = (c, slot);
            Enter(chunk, chunk.Keys[slot] == afterRow ? afterColumn + 1 : _firstColumn);
        }

        // Moves to the row at the current slot, before its first cell at or after the column.
        private void Enter(Chunk chunk, int column)
        {
            var number = chunk.Keys[_slot];
            if (number > _lastRow)
            {
                _row = null;
                return;
            }

            _rowNumber = number;
            _row = chunk.Rows[_slot];
            _cell = _row.LowerBound(column) - 1;
        }

        public void Reset()
        {
            _started = false;
            _row = null;
        }

        public readonly void Dispose()
        {
        }
    }

    /// <summary>Up to <see cref="ChunkSize"/> consecutive existing rows, in order.</summary>
    private sealed class Chunk
    {
        public readonly int[] Keys = new int[ChunkSize];
        public readonly RowData[] Rows = new RowData[ChunkSize];
        public int Count;

        public int LastKey => Keys[Count - 1];

        public void Append(int row, RowData data)
        {
            Keys[Count] = row;
            Rows[Count] = data;
            Count++;
        }

        public void Insert(int row, RowData data)
        {
            var index = LowerBound(row);
            Array.Copy(Keys, index, Keys, index + 1, Count - index);
            Array.Copy(Rows, index, Rows, index + 1, Count - index);
            Keys[index] = row;
            Rows[index] = data;
            Count++;
        }

        public void Remove(int row)
        {
            var index = LowerBound(row);
            Count--;
            Array.Copy(Keys, index + 1, Keys, index, Count - index);
            Array.Copy(Rows, index + 1, Rows, index, Count - index);
            Rows[Count] = null!;
        }

        /// <summary>Moves the upper half of the rows into a new chunk.</summary>
        public Chunk Split()
        {
            var upper = new Chunk { Count = Count / 2 };
            var kept = Count - upper.Count;
            Array.Copy(Keys, kept, upper.Keys, 0, upper.Count);
            Array.Copy(Rows, kept, upper.Rows, 0, upper.Count);
            Array.Clear(Rows, kept, upper.Count);
            Count = kept;
            return upper;
        }

        /// <summary>The index of the first row at or after <paramref name="row"/>; <see cref="Count"/> when there is none.</summary>
        public int LowerBound(int row)
        {
            int low = 0, high = Count - 1;
            while (low <= high)
            {
                var middle = (low + high) >>> 1;
                if (Keys[middle] < row)
                    low = middle + 1;
                else
                    high = middle - 1;
            }

            return low;
        }
    }

    private struct Entry(int column, CellData data)
    {
        public readonly int Column = column;
        public CellData Data = data;
    }

    /// <summary>The cells of one row, sorted by column.</summary>
    private sealed class RowData
    {
        public Entry[] Entries = [];
        public int Count;

        public RowData Copy(Func<CellData, CellData> copy)
        {
            var row = new RowData { Entries = new Entry[Count], Count = Count };
            for (var i = 0; i < Count; i++)
                row.Entries[i] = new Entry(Entries[i].Column, copy(Entries[i].Data));
            return row;
        }

        public CellData? Get(int column)
        {
            var index = Find(column);
            return index >= 0 ? Entries[index].Data : null;
        }

        public CellData GetOrCreate(int column, out bool created)
        {
            var index = Find(column);
            if (index >= 0)
            {
                created = false;
                return Entries[index].Data;
            }

            var cell = new CellData();
            Insert(~index, column, cell);
            created = true;
            return cell;
        }

        /// <summary>Puts the cell at the column; false when it replaced one there.</summary>
        public bool Put(int column, CellData data)
        {
            var index = Find(column);
            if (index >= 0)
            {
                Entries[index].Data = data;
                return false;
            }

            Insert(~index, column, data);
            return true;
        }

        public bool Remove(int column)
        {
            var index = Find(column);
            if (index < 0)
                return false;

            Count--;
            Array.Copy(Entries, index + 1, Entries, index, Count - index);
            Entries[Count] = default;
            return true;
        }

        public int LowerBound(int column)
        {
            var index = Find(column);
            return index >= 0 ? index : ~index;
        }

        /// <summary>The index of the column, or the bitwise complement of where it would go.</summary>
        private int Find(int column)
        {
            // Cells mostly come in column order: check the end first.
            var count = Count;
            if (count == 0 || Entries[count - 1].Column < column)
                return ~count;
            // In a row without gaps the column's index is its distance from the first column.
            var guess = column - Entries[0].Column;
            if ((uint)guess < (uint)count && Entries[guess].Column == column)
                return guess;
            int low = 0, high = count - 1;
            while (low <= high)
            {
                var middle = (low + high) >>> 1;
                var found = Entries[middle].Column;
                if (found == column)
                    return middle;
                if (found < column)
                    low = middle + 1;
                else
                    high = middle - 1;
            }

            return ~low;
        }

        private void Insert(int index, int column, CellData data)
        {
            if (Count == Entries.Length)
                Array.Resize(ref Entries, Math.Max(4, Entries.Length * 2));
            Array.Copy(Entries, index, Entries, index + 1, Count - index);
            Entries[index] = new Entry(column, data);
            Count++;
        }
    }
}
