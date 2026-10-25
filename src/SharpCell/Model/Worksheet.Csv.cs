using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SharpCell;

// Reading CSV text into a sheet.
public sealed partial class Worksheet
{
    /// <summary>
    /// Writes CSV text into the sheet starting at a cell, as Excel opens a .csv file: each field is
    /// read as if typed into its cell (see <see cref="Cell.Input"/>), so numbers, percentages, dates
    /// and logical values in the workbook's culture become those values, with a format taken from
    /// how they were written. Fields may be quoted, with quotes doubled and line breaks inside; rows
    /// end in CRLF, LF or CR. An empty field clears its cell. A field starting with <c>=</c> stays text
    /// unless <paramref name="formulas"/> is set, since a file from elsewhere could hide formulas in
    /// it. Nothing is written unless all of it fits on the sheet.
    /// </summary>
    /// <param name="topLeft">The cell of the first field, such as <c>A1</c>.</param>
    /// <param name="reader">The CSV text.</param>
    /// <param name="delimiter">What separates the fields; Excel uses <c>;</c> where the decimal separator is a comma.</param>
    /// <param name="formulas">Whether a field starting with <c>=</c> is a formula, as Excel reads it.</param>
    /// <returns>The range written: as many rows as the text has, as wide as its widest row.</returns>
    /// <exception cref="ArgumentException">The address is not a cell, or the delimiter is a quote or a line break.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The rows or fields run past the sheet's edge.</exception>
    /// <exception cref="FormulaParseException">With <paramref name="formulas"/>, a field starting with <c>=</c> is not a formula.</exception>
    public CellRange LoadCsv(string topLeft, TextReader reader, char delimiter = ',', bool formulas = false)
    {
        ArgumentNullException.ThrowIfNull(topLeft);
        ArgumentNullException.ThrowIfNull(reader);
        if (!CellAddress.TryParse(topLeft, out var start))
            throw new ArgumentException($"'{topLeft}' is not a cell address.", nameof(topLeft));
        if (delimiter is '"' or '\r' or '\n')
            throw new ArgumentException("A quote or a line break cannot separate fields.", nameof(delimiter));
        ThrowIfCannotChange();

        var rows = ReadCsv(reader, delimiter);
        var width = 0;
        foreach (var row in rows)
            width = Math.Max(width, row.Count);
        if (start.Row + rows.Count - 1 > CellAddress.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(reader), rows.Count, "The rows run past the last row of the sheet.");
        if (start.Column + width - 1 > CellAddress.MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(reader), width, "The fields run past the last column of the sheet.");

        for (var r = 0; r < rows.Count; r++)
        {
            for (var c = 0; c < rows[r].Count; c++)
            {
                var cell = new Cell(this, start.Row + r, start.Column + c);
                var field = rows[r][c];
                if (!formulas && field.StartsWith('='))
                    cell.Value = CellValue.Text(field);
                else
                    cell.Input = field;
            }
        }

        return new CellRange(this, new Evaluation.Area(start.Row, start.Column, start.Row + Math.Max(rows.Count, 1) - 1, start.Column + Math.Max(width, 1) - 1));
    }

    // RFC 4180, leniently: a quote inside an unquoted field is kept, text after a closing quote is
    // added to the field, and the last row may end without a line break.
    private static List<List<string>> ReadCsv(TextReader reader, char delimiter)
    {
        var rows = new List<List<string>>();
        var row = new List<string>();
        var field = new StringBuilder();
        var quoted = false;
        var fieldStarted = false;
        int next;
        while ((next = reader.Read()) >= 0)
        {
            var c = (char)next;
            if (quoted)
            {
                if (c != '"')
                    field.Append(c);
                else if (reader.Peek() == '"')
                    field.Append((char)reader.Read());
                else
                    quoted = false;
                continue;
            }

            if (c == '"' && !fieldStarted)
            {
                quoted = true;
                fieldStarted = true;
            }
            else if (c == delimiter)
            {
                row.Add(field.ToString());
                field.Clear();
                fieldStarted = false;
            }
            else if (c is '\r' or '\n')
            {
                if (c == '\r' && reader.Peek() == '\n')
                    reader.Read();
                row.Add(field.ToString());
                rows.Add(row);
                row = [];
                field.Clear();
                fieldStarted = false;
            }
            else
            {
                field.Append(c);
                fieldStarted = true;
            }
        }

        if (fieldStarted || field.Length > 0 || row.Count > 0)
        {
            row.Add(field.ToString());
            rows.Add(row);
        }

        return rows;
    }
}
