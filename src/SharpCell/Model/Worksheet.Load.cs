using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell;

// Filling a sheet from data.
public sealed partial class Worksheet
{
    /// <summary>
    /// Writes a header row and then one row per item of a collection, starting at a cell, and returns
    /// the range written (the header row included). The collection is enumerated once and nothing is
    /// written unless every value could be converted. Columns are lambdas, so no reflection is needed.
    /// </summary>
    /// <remarks>
    /// Values are converted like this:
    /// <list type="bullet">
    /// <item><description><c>null</c> and <see cref="DBNull"/> are empty and clear the cell; a <see cref="CellValue"/> is stored as it is.</description></item>
    /// <item><description>A <see cref="string"/> or <see cref="char"/> is text, never a formula: <c>"=1+1"</c> stays text.</description></item>
    /// <item><description>A <see cref="bool"/> is a logical value.</description></item>
    /// <item><description>Every integer type, <see cref="float"/>, <see cref="double"/> and <see cref="decimal"/> is a number; NaN and infinities are rejected.</description></item>
    /// <item><description>A <see cref="DateTime"/> is a date serial in the workbook's date system, shown in the culture's short date and, with a time of day, hours and minutes. <see cref="DateOnly"/> is a date at midnight; <see cref="DateTimeOffset"/> is the clock time it states.</description></item>
    /// <item><description>A <see cref="TimeOnly"/> is a fraction of a day with a time format; a <see cref="TimeSpan"/> is a number of days shown as <c>[h]:mm:ss</c>.</description></item>
    /// <item><description>An enum is its name; anything else is its <see cref="IFormattable"/> text in the workbook's culture, or its <c>ToString()</c>.</description></item>
    /// </list>
    /// Number cells keep the format they had; date and time cells take a format.
    /// </remarks>
    /// <typeparam name="T">The type of an item.</typeparam>
    /// <param name="topLeft">The cell of the first header, such as <c>B2</c>.</param>
    /// <param name="rows">The items; one row each.</param>
    /// <param name="columns">The header of each column and how to get its value from an item.</param>
    /// <returns>The range written, header row included.</returns>
    /// <exception cref="ArgumentNullException">An argument or a header is null.</exception>
    /// <exception cref="ArgumentException">There are no columns, the address is not a cell, or a value cannot be stored (NaN, infinity).</exception>
    /// <exception cref="ArgumentOutOfRangeException">The data runs past the last row or column, or a date is before the date system starts.</exception>
    public CellRange LoadRows<T>(string topLeft, IEnumerable<T> rows, params (string Header, Func<T, object?> Value)[] columns)
    {
        ArgumentNullException.ThrowIfNull(topLeft);
        ArgumentNullException.ThrowIfNull(rows);
        ArgumentNullException.ThrowIfNull(columns);
        if (columns.Length == 0)
            throw new ArgumentException("At least one column is needed.", nameof(columns));
        foreach (var (name, value) in columns)
        {
            if (name is null)
                throw new ArgumentNullException(nameof(columns), "A column header is null.");
            if (value is null)
                throw new ArgumentNullException(nameof(columns), $"The column '{name}' has no value function.");
        }

        if (!CellAddress.TryParse(topLeft, out var start))
            throw new ArgumentException($"'{topLeft}' is not a cell address.", nameof(topLeft));
        if (start.Column + columns.Length - 1 > CellAddress.MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(columns), columns.Length, "The columns run past the last column of the sheet.");
        Workbook.ThrowIfInCustomFunction();

        // Convert everything first, so that a value that cannot be stored leaves the sheet as it was.
        var converted = new List<(CellValue Value, string? Format)[]>();
        var header = new (CellValue, string?)[columns.Length];
        for (var c = 0; c < columns.Length; c++)
            header[c] = (CellValue.Text(columns[c].Header), null);
        converted.Add(header);

        foreach (var item in rows)
        {
            var row = start.Row + converted.Count;
            if (row > CellAddress.MaxRow)
                throw new ArgumentOutOfRangeException(nameof(rows), "The rows run past the last row of the sheet.");
            var line = new (CellValue, string?)[columns.Length];
            for (var c = 0; c < columns.Length; c++)
            {
                try
                {
                    line[c] = (ObjectValues.Convert(columns[c].Value(item), Workbook, out var format), format);
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    throw new ArgumentOutOfRangeException(nameof(rows), ex.ActualValue, $"{new CellAddress(row, start.Column + c)}: {ex.Message}");
                }
                catch (ArgumentException ex)
                {
                    throw new ArgumentException($"{new CellAddress(row, start.Column + c)}: {ex.Message}", nameof(rows), ex);
                }
            }

            converted.Add(line);
        }

        var values = new CellValue[converted.Count, columns.Length];
        for (var r = 0; r < converted.Count; r++)
        {
            for (var c = 0; c < columns.Length; c++)
                values[r, c] = converted[r][c].Value;
        }

        var range = new CellRange(this, new Area(start.Row, start.Column, start.Row + converted.Count - 1, start.Column + columns.Length - 1));
        range.SetValues(values);
        for (var r = 1; r < converted.Count; r++)
        {
            for (var c = 0; c < columns.Length; c++)
            {
                if (converted[r][c].Format is { } format)
                    this[start.Row + r, start.Column + c].NumberFormat = format;
            }
        }

        return range;
    }
}
