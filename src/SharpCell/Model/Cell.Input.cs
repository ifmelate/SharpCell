using System;

namespace SharpCell;

public sealed partial class Cell
{
    /// <summary>
    /// The cell as typed in Excel. Setting it does what typing the text and pressing Enter does:
    /// <c>=</c> starts a formula; <c>'</c> keeps the rest as text; <c>TRUE</c>, <c>FALSE</c> and error
    /// literals such as <c>#N/A</c> are those values; numbers, percentages, currency and grouped
    /// digits, dates and times are read in the workbook's <see cref="Workbook.Culture"/>, as Excel
    /// reads them in arithmetic and DATEVALUE (a date without a year is in the current year);
    /// anything else is text; an empty string clears the cell's content. A cell in the General
    /// format takes a format from what was typed, as in Excel: <c>12%</c> gives <c>0%</c>,
    /// <c>1,250.50</c> gives <c>#,##0.00</c>, <c>$5</c> a currency format, a date the culture's
    /// short date, a time <c>h:mm</c> with seconds and AM/PM when they were typed. A cell with a
    /// format of its own keeps it.
    /// <para>
    /// Getting it gives the text Excel shows to edit the cell, which set again gives the same
    /// content: the formula, a number in full (a percentage as one, a date or time as the culture
    /// writes it), text with an apostrophe in front when it would read as something else.
    /// </para>
    /// </summary>
    /// <exception cref="FormulaParseException">The text starts with <c>=</c> but is not a formula; the cell is unchanged.</exception>
    /// <exception cref="InvalidOperationException">The cell is part of an array formula other than its top-left cell.</exception>
    public string Input
    {
        get => CellInput.EditText(this);
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            var workbook = Worksheet.Workbook;
            var entry = CellInput.Read(value, workbook.Culture, workbook.DateSystem, workbook.Clock.GetLocalNow().Year);
            if (entry.Formula is { } formula)
                Formula = formula;
            else
                Value = entry.Value;
            if (entry.NumberFormat is { } format && Worksheet.FormatAt(Row, Column) is null)
                NumberFormat = format;
        }
    }
}
