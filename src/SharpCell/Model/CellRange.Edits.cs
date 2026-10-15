using System;

namespace SharpCell;

/// <summary>Where the cells beside inserted cells go, as in Excel's Insert dialog.</summary>
public enum InsertShift
{
    /// <summary>The cells below move down: "Shift cells down".</summary>
    Down,

    /// <summary>The cells to the right move right: "Shift cells right".</summary>
    Right,
}

/// <summary>Which cells take the place of deleted cells, as in Excel's Delete dialog.</summary>
public enum DeleteShift
{
    /// <summary>The cells below move up: "Shift cells up".</summary>
    Up,

    /// <summary>The cells to the right move left: "Shift cells left".</summary>
    Left,
}

// Inserting and deleting cells, as Excel's Insert and Delete with a shift do.
public sealed partial class CellRange
{
    /// <summary>
    /// Inserts empty cells in place of the range, as Excel's Insert with "Shift cells down" or
    /// "Shift cells right" does: only the cells in the range's columns (or rows) move, along with
    /// their styles, number formats, hyperlinks, notes, data validation, merged areas and tables, and every
    /// reference to them follows. A reference to a range that only some of the moved cells are in
    /// stays as it was, as in Excel; data validation over such a range is split. A range of whole
    /// rows (or whole columns) inserts rows (or columns), as <see cref="Worksheet.InsertRows"/> does.
    /// Every formula is calculated again at the next <see cref="Workbook.Recalculate"/>.
    /// </summary>
    /// <param name="shift">Where the cells beside the range go.</param>
    /// <exception cref="ArgumentException"><paramref name="shift"/> is not a defined value.</exception>
    /// <exception cref="InvalidOperationException">Content would be pushed off the sheet, or the cells that
    /// move would split a merged area, a table or an array formula.</exception>
    /// <exception cref="NotSupportedException">A formula on the sheet, or one naming it, cannot be read by
    /// SharpCell, so its references could not follow; or cells would be inserted inside a table's columns.</exception>
    public void Insert(InsertShift shift)
    {
        if (shift is not (InsertShift.Down or InsertShift.Right))
            throw new ArgumentException($"{shift} is not a way to insert cells.", nameof(shift));
        Worksheet.EditCells(Area, rows: shift == InsertShift.Down, delete: false);
    }

    /// <summary>
    /// Deletes the range's cells, as Excel's Delete with "Shift cells up" or "Shift cells left"
    /// does: the cells below (or to the right) in the range's columns (or rows) move into their
    /// place. A reference to a deleted cell becomes <c>#REF!</c>; otherwise references follow as
    /// with <see cref="Insert"/>. A range of whole rows (or columns) deletes them.
    /// </summary>
    /// <param name="shift">Which cells take the place of the deleted ones.</param>
    /// <exception cref="ArgumentException"><paramref name="shift"/> is not a defined value.</exception>
    /// <exception cref="InvalidOperationException">The cells that move would split a merged area, a table or an
    /// array formula, or a table would lose its header, totals or data rows.</exception>
    /// <exception cref="NotSupportedException">A formula on the sheet, or one naming it, cannot be read by
    /// SharpCell; or some of a table's columns would be deleted.</exception>
    public void Delete(DeleteShift shift)
    {
        if (shift is not (DeleteShift.Up or DeleteShift.Left))
            throw new ArgumentException($"{shift} is not a way to delete cells.", nameof(shift));
        Worksheet.EditCells(Area, rows: shift == DeleteShift.Up, delete: true);
    }
}
