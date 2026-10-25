using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using SharpCell.Evaluation;
using SharpCell.Functions;
using SharpCell.Parsing;

namespace SharpCell;

/// <summary>A key of <see cref="CellRange.Sort(SortKey[])"/>: the column rows are put in order by.</summary>
/// <param name="Index">The sheet column (A is 1) whose cells order the rows; for a sort left to right, the sheet row ordering the columns.</param>
/// <param name="Descending">Whether the highest come first.</param>
public readonly record struct SortKey(int Index, bool Descending = false);

// Sorting cells, as Excel's Data > Sort does.
public sealed partial class CellRange
{
    /// <summary>
    /// Sorts the rows of the range, as Excel's Data &gt; Sort does: by the first key's column, rows
    /// equal there by the second's, and so on; rows equal by every key keep their order. Without a
    /// key the first column orders the rows, ascending. See <see cref="Sort(IReadOnlyList{SortKey}, bool, bool)"/>.
    /// </summary>
    /// <param name="keys">The columns that order the rows.</param>
    /// <exception cref="ArgumentOutOfRangeException">A key's column is outside the range.</exception>
    /// <exception cref="InvalidOperationException">The range holds merged cells, or part of an array formula or a spill.</exception>
    /// <exception cref="NotSupportedException">The range holds a formula SharpCell cannot read (such as a link to another workbook).</exception>
    public void Sort(params SortKey[] keys) => Sort(keys, matchCase: false, leftToRight: false);

    /// <summary>
    /// Sorts the rows of the range, or its columns, as Excel's Data &gt; Sort does. Ascending,
    /// numbers come first, then text (ignoring case unless <paramref name="matchCase"/>), then FALSE
    /// and TRUE, then errors, which are all equal; descending reverses that, and empty cells come last
    /// either way. Formulas out of date are calculated first.
    /// <para>
    /// The cells move with what they hold: values, formulas, number formats, styles, hyperlinks and
    /// notes. A moved formula's relative references keep their distance from its cell, as when it is
    /// copied, so <c>=B2*C2</c> moved to row 5 becomes <c>=B5*C5</c>; formulas elsewhere that point
    /// into the range are not changed, and read whatever was sorted into their cells. Data
    /// validation, conditional formats, row heights and hidden rows stay where they are. Every formula
    /// reading the cells is calculated again at the next <see cref="Workbook.Recalculate"/>.
    /// </para>
    /// </summary>
    /// <param name="keys">The columns that order the rows (or the rows that order the columns), first key first.</param>
    /// <param name="matchCase">Whether text of different case differs, lower case first, as Excel's Case sensitive option.</param>
    /// <param name="leftToRight">Whether to sort the columns by rows, as Excel's Sort left to right option.</param>
    /// <exception cref="ArgumentOutOfRangeException">A key's column (or row) is outside the range.</exception>
    /// <exception cref="InvalidOperationException">The range holds merged cells, or part of an array formula or a spill.</exception>
    /// <exception cref="NotSupportedException">The range holds a formula SharpCell cannot read (such as a link to another workbook).</exception>
    public void Sort(IReadOnlyList<SortKey> keys, bool matchCase = false, bool leftToRight = false)
    {
        ArgumentNullException.ThrowIfNull(keys);
        Worksheet.ThrowIfCannotChange();
        var area = Used();
        if (area is not { } sorted)
            return;
        var (first, last) = leftToRight ? (sorted.FirstRow, sorted.LastRow) : (sorted.FirstColumn, sorted.LastColumn);
        foreach (var key in keys)
        {
            if (key.Index < first || key.Index > last)
                throw new ArgumentOutOfRangeException(nameof(keys), key.Index, leftToRight ? "A key's row is outside the range." : "A key's column is outside the range.");
        }

        CheckSortable(sorted);
        var order = Order(sorted, keys.Count > 0 ? keys : [new SortKey(first)], matchCase, leftToRight);
        var moved = false;
        for (var i = 0; i < order.Length && !moved; i++)
            moved = order[i] != i;
        if (moved)
            Permute(sorted, order, leftToRight);
    }

    // The range down to the rows and columns used: sorting A:C sorts the rows that hold something.
    private Area? Used()
    {
        if (Worksheet.Store.Bounds is not { } used || !Area.TryIntersect(used, out var area))
            return null;
        return area;
    }

    private void CheckSortable(Area area)
    {
        foreach (var merged in Worksheet.MergedAreaList)
        {
            if (merged.TryIntersect(area, out _))
                throw new InvalidOperationException($"The range holds the merged cells {new CellRange(Worksheet, merged).Address}; unmerge them to sort it.");
        }

        var calculation = Worksheet.Workbook.Calculation;
        foreach (var cell in Worksheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
        {
            var data = cell.Data;
            var at = new Cell(Worksheet, cell.Row, cell.Column);
            if (data.FixedArray is not null || calculation.FixedArrayAt(Worksheet, cell.Row, cell.Column) is not null)
                throw new InvalidOperationException($"{at} is part of an array formula, which cannot be sorted.");
            if (data.SpillAnchor is not null || data.SpillArea is { IsSingleCell: false })
                throw new InvalidOperationException($"{at} is part of a spill, which cannot be sorted.");
            if (data.Formula is UnsupportedNode)
                throw new NotSupportedException($"{at} has a formula SharpCell cannot read; its references could not be moved.");
        }
    }

    // Where each line (row, or column left to right) goes: order[i] is the line that comes i-th.
    private int[] Order(Area area, IReadOnlyList<SortKey> keys, bool matchCase, bool leftToRight)
    {
        var (start, count) = leftToRight ? (area.FirstColumn, area.Columns) : (area.FirstRow, area.Rows);
        var values = new CellValue[keys.Count][];
        for (var k = 0; k < keys.Count; k++)
        {
            values[k] = new CellValue[count];
            for (var i = 0; i < count; i++)
            {
                values[k][i] = leftToRight
                    ? Worksheet.CurrentValue(keys[k].Index, start + i)
                    : Worksheet.CurrentValue(start + i, keys[k].Index);
            }
        }

        var culture = Worksheet.Workbook.Culture;
        var order = Enumerable.Range(0, count).ToArray();
        // A stable sort: lines equal by every key keep their order.
        Array.Sort(order, (a, b) =>
        {
            for (var k = 0; k < keys.Count; k++)
            {
                var compared = Compare(values[k][a], values[k][b], keys[k].Descending, matchCase, culture);
                if (compared != 0)
                    return compared;
            }

            return a.CompareTo(b);
        });
        return order;
    }

    // Data > Sort's order: as SORT's, but all errors are equal, and case may count.
    private static int Compare(CellValue a, CellValue b, bool descending, bool matchCase, CultureInfo culture)
    {
        if (a.IsError && b.IsError)
            return 0;
        var order = ArrayShapeSortFunctions.CompareForSort(a, b, descending, culture);
        if (order != 0 || !matchCase || a.Kind != CellValueKind.Text || b.Kind != CellValueKind.Text)
            return order;
        order = culture.CompareInfo.Compare(a.AsText(), b.AsText(), CompareOptions.None);
        return descending ? -order : order;
    }

    private void Permute(Area area, int[] order, bool leftToRight)
    {
        // Where the line at each offset goes.
        var to = new int[order.Length];
        for (var i = 0; i < order.Length; i++)
            to[order[i]] = i;
        (int Row, int Column) Moved((int Row, int Column) offset) =>
            leftToRight ? (offset.Row, to[offset.Column]) : (to[offset.Row], offset.Column);

        var cells = new List<Copied>();
        foreach (var cell in Worksheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
        {
            var (row, column) = Moved((cell.Row - area.FirstRow, cell.Column - area.FirstColumn));
            cells.Add(new Copied(row, column, cell.Data.Value, cell.Data.Formula, cell.Data.IsLegacy));
        }

        var positions = cells.Select(cell => (cell.Row, cell.Column)).ToHashSet();
        var styles = Within(Worksheet.Styles, area).ToDictionary(pair => Moved(pair.Key), pair => pair.Value);
        var formats = Within(Worksheet.Formats, area).ToDictionary(pair => Moved(pair.Key), pair => pair.Value);
        var links = Within(Worksheet.Hyperlinks, area).ToDictionary(pair => Moved(pair.Key), pair => pair.Value);
        var notes = Within(Worksheet.NoteMap, area).ToDictionary(pair => Moved(pair.Key), pair => pair.Value);

        Paste(Worksheet, area, cells, positions, styles, formats);
        Replace(Worksheet.Hyperlinks, links, area, Worksheet.SetHyperlink);
        Replace(Worksheet.NoteMap, notes, area, Worksheet.SetNote);
    }
}
