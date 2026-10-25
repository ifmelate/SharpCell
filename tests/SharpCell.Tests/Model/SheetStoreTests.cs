using System;
using System.Collections.Generic;
using System.Linq;
using SharpCell.Evaluation;

namespace SharpCell.Tests.Model;

/// <summary>The sparse cell store, checked against a plain sorted dictionary of the same cells.</summary>
public class SheetStoreTests
{
    private sealed class Model
    {
        public SheetStore Store { get; } = new();

        public SortedDictionary<(int Row, int Column), CellData> Cells { get; } = [];

        public void Add(int row, int column)
        {
            var data = Store.GetOrCreate(row, column);
            if (Cells.TryGetValue((row, column), out var existing))
                Assert.Same(existing, data);
            else
                Cells.Add((row, column), data);
        }

        public void Remove(int row, int column)
        {
            Store.Remove(row, column);
            Cells.Remove((row, column));
        }

        public void Check(int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            var expected = Cells
                .Where(c => c.Key.Row >= firstRow && c.Key.Row <= lastRow && c.Key.Column >= firstColumn && c.Key.Column <= lastColumn)
                .Select(c => (c.Key.Row, c.Key.Column, c.Value))
                .ToArray();
            var actual = Store.Enumerate(firstRow, firstColumn, lastRow, lastColumn).Select(c => (c.Row, c.Column, c.Data)).ToArray();
            Assert.Equal(expected.Length, actual.Length);
            for (var i = 0; i < expected.Length; i++)
            {
                Assert.Equal((expected[i].Row, expected[i].Column), (actual[i].Row, actual[i].Column));
                Assert.Same(expected[i].Value, actual[i].Data);
            }
        }

        public void CheckAll()
        {
            Assert.Equal(Cells.Count, Store.Count);
            Check(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn);
            if (Cells.Count == 0)
            {
                Assert.Null(Store.Bounds);
            }
            else
            {
                var bounds = new Area(Cells.Keys.Min(k => k.Row), Cells.Keys.Min(k => k.Column),
                    Cells.Keys.Max(k => k.Row), Cells.Keys.Max(k => k.Column));
                Assert.Equal(bounds, Store.Bounds);
            }

            foreach (var ((row, column), data) in Cells)
                Assert.Same(data, Store.Get(row, column));
        }
    }

    [Fact]
    public void An_empty_store_has_no_cells_and_no_bounds()
    {
        var store = new SheetStore();
        Assert.Empty(store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn));
        Assert.Null(store.Bounds);
        Assert.Null(store.Get(1, 1));
        store.Remove(1, 1);
        Assert.Equal(0, store.Count);
    }

    [Fact]
    public void Rows_created_out_of_order_are_enumerated_in_order()
    {
        var model = new Model();
        var random = new Random(7);
        foreach (var row in Enumerable.Range(1, 3000).OrderBy(_ => random.Next()))
            model.Add(row * 3, 1 + row % 5);
        for (var row = 9000; row >= 1; row -= 7)
            model.Add(row, 2);
        model.CheckAll();
        model.Check(100, 1, 200, 3);
        model.Check(4500, 2, 4500, 2);
        model.Check(8999, 1, CellAddress.MaxRow, CellAddress.MaxColumn);
    }

    [Fact]
    public void Enumeration_keeps_to_the_rectangle_over_dense_rows()
    {
        var model = new Model();
        for (var row = 1; row <= 2000; row++)
        {
            for (var column = 1; column <= 9; column++)
                model.Add(row, column);
        }

        model.Check(1, 4, CellAddress.MaxRow, 4);
        model.Check(2, 1, 2001, 9);
        model.Check(500, 3, 1500, 5);
        model.Check(1, 10, CellAddress.MaxRow, CellAddress.MaxColumn);
        model.Check(2001, 1, CellAddress.MaxRow, 9);
        model.Check(1, 1, 1, 1);
        model.Check(2000, 9, 2000, 9);
    }

    [Fact]
    public void Removing_the_last_cell_of_a_row_drops_the_row()
    {
        var model = new Model();
        model.Add(5, 3);
        model.Add(7, 2);
        model.Add(7, 9);
        model.Remove(5, 3);
        model.CheckAll();
        Assert.Equal(new Area(7, 2, 7, 9), model.Store.Bounds);
        model.Remove(7, 9);
        model.Remove(7, 2);
        model.CheckAll();
        model.Add(6, 1);
        model.CheckAll();
    }

    [Fact]
    public void Random_adds_and_removes_match_a_sorted_dictionary()
    {
        var model = new Model();
        var random = new Random(42);
        for (var step = 0; step < 20_000; step++)
        {
            var row = random.Next(1, 5000);
            var column = random.Next(1, 12);
            if (random.Next(3) == 0)
                model.Remove(row, column);
            else
                model.Add(row, column);
            if (step % 2000 == 0)
            {
                model.CheckAll();
                var top = random.Next(1, 5000);
                var left = random.Next(1, 12);
                model.Check(top, left, top + random.Next(0, 2000), left + random.Next(0, 4));
            }
        }

        model.CheckAll();
        foreach (var key in model.Cells.Keys.ToArray())
            model.Remove(key.Row, key.Column);
        model.CheckAll();
    }

    [Fact]
    public void Move_sends_cells_where_the_map_says_and_drops_the_rest()
    {
        var model = new Model();
        for (var row = 1; row <= 600; row++)
        {
            model.Add(row, 1);
            model.Add(row, 4);
        }

        var before = model.Cells.ToDictionary(c => c.Key, c => c.Value);
        // Reverse the rows, drop column 4 of even rows.
        model.Store.Move((row, column) => column == 4 && row % 2 == 0 ? null : new CellAddress(601 - row, column));
        model.Cells.Clear();
        foreach (var ((row, column), data) in before)
        {
            if (!(column == 4 && row % 2 == 0))
                model.Cells.Add((601 - row, column), data);
        }

        model.CheckAll();
        model.Add(1000, 1);
        model.Remove(1, 1);
        model.CheckAll();
    }

    [Fact]
    public void A_copy_is_independent_of_its_source()
    {
        var source = new Model();
        for (var row = 1; row <= 700; row++)
        {
            source.Add(row * 2, 1);
            source.Add(row * 2, 3);
        }

        var copy = new SheetStore();
        var copies = new Dictionary<CellData, CellData>();
        copy.CopyFrom(source.Store, data =>
        {
            var copied = new CellData { Value = data.Value };
            copies.Add(data, copied);
            return copied;
        });

        Assert.Equal(source.Store.Count, copy.Count);
        Assert.Equal(source.Store.Bounds, copy.Bounds);
        Assert.Equal(
            source.Cells.Select(c => (c.Key.Row, c.Key.Column, copies[c.Value])),
            copy.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn).Select(c => (c.Row, c.Column, c.Data)));

        source.Remove(2, 1);
        source.Add(3, 1);
        source.Add(5000, 7);
        source.CheckAll();
        Assert.NotNull(copy.Get(2, 1));
        Assert.Null(copy.Get(3, 1));
        Assert.Equal(1400, copy.Count);
        copy.Remove(4, 3);
        copy.GetOrCreate(1, 1);
        Assert.Same(source.Cells[(4, 3)], source.Store.Get(4, 3));
        Assert.Null(source.Store.Get(1, 1));
    }

    [Fact]
    public void Changes_during_an_enumeration_never_repeat_or_skip_cells_that_stay()
    {
        var store = new SheetStore();
        for (var row = 1; row <= 1000; row++)
        {
            store.GetOrCreate(row * 2, 1);
            store.GetOrCreate(row * 2, 2);
        }

        var seen = new List<(int Row, int Column)>();
        var removed = new HashSet<(int Row, int Column)>();
        foreach (var cell in store.Enumerate(1, 1, CellAddress.MaxRow, 2))
        {
            seen.Add((cell.Row, cell.Column));
            // Cells added before, after and in the current row, and rows removed ahead.
            store.GetOrCreate(cell.Row - 1, 1);
            if (cell.Column == 1)
                store.GetOrCreate(cell.Row, 3);
            var ahead = cell.Row + 10;
            if (ahead % 2 == 0 && removed.Add((ahead, 1)))
            {
                store.Remove(ahead, 1);
                store.Remove(ahead, 2);
                removed.Add((ahead, 2));
            }
        }

        for (var i = 1; i < seen.Count; i++)
            Assert.True(seen[i - 1].Row < seen[i].Row || (seen[i - 1].Row == seen[i].Row && seen[i - 1].Column < seen[i].Column));
        for (var row = 1; row <= 1000; row++)
        {
            for (var column = 1; column <= 2; column++)
            {
                if (!removed.Contains((row * 2, column)))
                    Assert.Contains((row * 2, column), seen);
            }
        }
    }
}
