using System;
using System.Collections.Generic;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell;

/// <summary>
/// A rectangle of cells on one sheet, such as <c>A1:C10</c>. Like a <see cref="Cell"/>, it is a
/// handle: creating one allocates no cells.
/// </summary>
public sealed class CellRange : IEquatable<CellRange>
{
    internal CellRange(Worksheet worksheet, Area area)
    {
        Worksheet = worksheet;
        Area = area;
    }

    /// <summary>The sheet the range is on.</summary>
    public Worksheet Worksheet { get; }

    /// <summary>The first row, from 1.</summary>
    public int FirstRow => Area.FirstRow;

    /// <summary>The first column, from 1 (A is 1).</summary>
    public int FirstColumn => Area.FirstColumn;

    /// <summary>The last row, from 1.</summary>
    public int LastRow => Area.LastRow;

    /// <summary>The last column, from 1.</summary>
    public int LastColumn => Area.LastColumn;

    /// <summary>The A1 address without the sheet name: <c>A1:C10</c>, or <c>B2</c> for one cell.</summary>
    public string Address => Area.IsSingleCell
        ? new CellAddress(FirstRow, FirstColumn).ToString()
        : $"{new CellAddress(FirstRow, FirstColumn)}:{new CellAddress(LastRow, LastColumn)}";

    internal Area Area { get; }

    /// <summary>
    /// The cells of the range that hold something (a constant, a formula or a value spilled into
    /// them), row by row and left to right. Only existing cells are visited, so <c>A:A</c> is cheap.
    /// Values and formulas of the cells may be changed while enumerating; adding or removing a cell
    /// of the sheet makes the next step throw <see cref="InvalidOperationException"/>.
    /// </summary>
    public IEnumerable<Cell> Cells => EnumerateCells(Worksheet, Area);

    /// <summary>
    /// The value of every cell, empty ones included, indexed <c>[row, column]</c> from the range's
    /// top-left cell. A whole column allocates an element for each of its 1,048,576 cells; use
    /// <see cref="Worksheet.UsedRange"/> or <see cref="Cells"/> for sparse data.
    /// </summary>
    /// <exception cref="InvalidOperationException">The range has more cells than an array can hold.</exception>
    public CellValue[,] GetValues()
    {
        if (Area.CellCount > Array.MaxLength)
            throw new InvalidOperationException($"{this} has {Area.CellCount} cells, more than one array can hold.");

        var values = new CellValue[Area.Rows, Area.Columns];
        foreach (var cell in Worksheet.Store.Enumerate(FirstRow, FirstColumn, LastRow, LastColumn))
            values[cell.Row - FirstRow, cell.Column - FirstColumn] = cell.Data.Value;
        return values;
    }

    /// <summary>
    /// Sets the value of every cell, as assigning <see cref="Cell.Value"/> would: formulas are
    /// replaced and <see cref="CellValue.Empty"/> clears a cell. Everything is checked before the
    /// first cell is written, so an exception leaves the sheet unchanged; a cell of an array formula
    /// other than its top-left cell is refused even when the range covers the whole array.
    /// </summary>
    /// <param name="values">Values indexed <c>[row, column]</c>, exactly as large as the range.</param>
    /// <exception cref="ArgumentException">The array's size differs from the range's, or it holds a
    /// Missing, array or lambda value.</exception>
    /// <exception cref="InvalidOperationException">A cell is part of an array formula other than its top-left cell.</exception>
    public void SetValues(CellValue[,] values)
    {
        ArgumentNullException.ThrowIfNull(values);
        Worksheet.Workbook.ThrowIfInCustomFunction();
        if (values.GetLength(0) != Area.Rows || values.GetLength(1) != Area.Columns)
            throw new ArgumentException(
                $"The array is {values.GetLength(0)}x{values.GetLength(1)}; {this} is {Area.Rows}x{Area.Columns}.", nameof(values));
        foreach (var value in values)
        {
            if (value.Kind is CellValueKind.Missing or CellValueKind.Array or CellValueKind.Lambda)
                throw new ArgumentException($"A cell cannot hold a {value.Kind} value.", nameof(values));
        }

        var calculation = Worksheet.Workbook.Calculation;
        for (var row = FirstRow; row <= LastRow; row++)
        {
            for (var column = FirstColumn; column <= LastColumn; column++)
            {
                if (calculation.FixedArrayAt(Worksheet, row, column) is { } anchor)
                    throw new InvalidOperationException($"{new Cell(Worksheet, row, column)} is part of the array formula at {anchor}; change the whole array there.");
            }
        }

        for (var r = 0; r < Area.Rows; r++)
        {
            for (var c = 0; c < Area.Columns; c++)
            {
                var value = values[r, c];
                if (value.Kind == CellValueKind.Empty && Worksheet.Store.Get(FirstRow + r, FirstColumn + c) is null)
                    continue;
                new Cell(Worksheet, FirstRow + r, FirstColumn + c).Value = value;
            }
        }
    }

    // Checks the store's version before every step, so the underlying enumeration never runs on a
    // changed store.
    internal static IEnumerable<Cell> EnumerateCells(Worksheet sheet, Area area)
    {
        var store = sheet.Store;
        var version = store.Version;
        using var cells = store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn).GetEnumerator();
        while (true)
        {
            if (store.Version != version)
                throw new InvalidOperationException($"Cells of {sheet.Name} were added or removed during the enumeration.");
            if (!cells.MoveNext())
                yield break;
            yield return new Cell(sheet, cells.Current.Row, cells.Current.Column);
        }
    }

    /// <summary>Whether the other range is on the same sheet with the same corners.</summary>
    public bool Equals(CellRange? other) => other is not null && other.Worksheet == Worksheet && other.Area == Area;

    /// <summary>Whether <paramref name="obj"/> is a range on the same sheet with the same corners.</summary>
    public override bool Equals(object? obj) => Equals(obj as CellRange);

    /// <summary>A hash code consistent with <see cref="Equals(CellRange)"/>.</summary>
    public override int GetHashCode() => HashCode.Combine(Worksheet, Area);

    /// <summary>The address with the sheet name as a formula writes it, such as <c>Sheet1!A1:C10</c> or <c>'My Sheet'!B2</c>.</summary>
    public override string ToString() => FormulaPrinter.QuoteSheetName(Worksheet.Name) + "!" + Address;
}
