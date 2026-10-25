using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using SharpCell.Evaluation;

namespace SharpCell;

/// <summary>A row of a range read by <see cref="CellRange.ReadRows{T}"/>: its cells by header or by column.</summary>
public readonly struct RangeRow
{
    private readonly Worksheet _sheet;
    private readonly IReadOnlyDictionary<string, int> _columns;

    internal RangeRow(Worksheet sheet, int row, IReadOnlyDictionary<string, int> columns)
    {
        _sheet = sheet;
        Row = row;
        _columns = columns;
    }

    /// <summary>The sheet row, from 1.</summary>
    public int Row { get; }

    /// <summary>The cell under a header of the range's first row, ignoring case.</summary>
    /// <exception cref="KeyNotFoundException">No column has that header.</exception>
    public Cell this[string header] =>
        _columns.TryGetValue(header ?? throw new ArgumentNullException(nameof(header)), out var column)
            ? new Cell(_sheet, Row, column)
            : throw new KeyNotFoundException($"No column has the header '{header}'.");

    /// <summary>The cell in a sheet column (A is 1).</summary>
    public Cell this[int column] => _sheet[Row, column];

    /// <summary>Whether a column has the header, ignoring case.</summary>
    public bool HasColumn(string header) => _columns.ContainsKey(header);
}

// Reading a range out: as a DataTable, as objects, as CSV text.
public sealed partial class CellRange
{
    /// <summary>
    /// The range as a <see cref="DataTable"/> named after the sheet, one row per row of the range.
    /// With <paramref name="headers"/>, the first row names the columns (<c>ColumnN</c> for an empty
    /// one, a number added to a repeated one, as tables name them); otherwise they are <c>Column1</c>,
    /// <c>Column2</c> and so on. A column whose values are all numbers is <see cref="double"/>, all
    /// shown as dates <see cref="DateTime"/>, all logical values <see cref="bool"/>, all text
    /// <see cref="string"/>; any other mix is <see cref="object"/>, with errors as their text
    /// (<c>#N/A</c>). Empty cells are <see cref="DBNull"/>. Whole columns such as <c>A:C</c> end at the
    /// last row used. Formulas out of date are calculated first.
    /// </summary>
    /// <param name="headers">Whether the first row holds the column names.</param>
    public DataTable ToDataTable(bool headers = true)
    {
        var area = Exported();
        var table = new DataTable(Worksheet.Name) { Locale = Worksheet.Workbook.Culture };
        var firstData = area.FirstRow + (headers ? 1 : 0);
        var names = ColumnNames(area, headers);
        var system = Worksheet.Workbook.DateSystem;
        for (var c = 0; c < area.Columns; c++)
        {
            var column = area.FirstColumn + c;
            var kinds = new HashSet<CellValueKind>();
            var dates = true;
            for (var row = firstData; row <= area.LastRow; row++)
            {
                var value = Worksheet.CurrentValue(row, column);
                if (value.Kind == CellValueKind.Empty)
                    continue;
                kinds.Add(value.Kind);
                dates &= value.Kind == CellValueKind.Number && Worksheet.FormatAt(row, column) is { ShowsDate: true }
                    && DateSerial.TryToDateTime(value.AsNumber(), system, out _);
            }

            var type = kinds.Count != 1 ? typeof(object) : kinds.Single() switch
            {
                CellValueKind.Number => dates ? typeof(DateTime) : typeof(double),
                CellValueKind.Boolean => typeof(bool),
                CellValueKind.Text => typeof(string),
                _ => typeof(object),
            };
            table.Columns.Add(names[c], type);
        }

        for (var row = firstData; row <= area.LastRow; row++)
        {
            var items = new object[area.Columns];
            for (var c = 0; c < area.Columns; c++)
            {
                var value = Worksheet.CurrentValue(row, area.FirstColumn + c);
                items[c] = value.Kind switch
                {
                    CellValueKind.Empty => DBNull.Value,
                    CellValueKind.Number when table.Columns[c].DataType == typeof(DateTime) => value.AsDateTime(system),
                    CellValueKind.Number => value.AsNumber(),
                    CellValueKind.Boolean => value.AsBoolean(),
                    CellValueKind.Text => value.AsText(),
                    _ => value.AsError().ToText(),
                };
            }

            table.Rows.Add(items);
        }

        table.AcceptChanges();
        return table;
    }

    /// <summary>
    /// Reads each row below the range's first row into an object, the first row naming the
    /// columns: <c>range.ReadRows(row =&gt; new Order(row["Region"].Text, row["Units"].Value.AsNumber()))</c>.
    /// The lambda picks the cells, so no reflection is needed. Whole columns such as <c>A:C</c> end at
    /// the last row used. Formulas out of date in the range are calculated first, so the cells hold
    /// their current values.
    /// </summary>
    /// <typeparam name="T">What a row becomes.</typeparam>
    /// <param name="read">Makes an item from a row.</param>
    /// <returns>The items, in row order.</returns>
    public List<T> ReadRows<T>(Func<RangeRow, T> read)
    {
        ArgumentNullException.ThrowIfNull(read);
        var area = Exported();
        var names = ColumnNames(area, headers: true);
        var columns = new Dictionary<string, int>(StringComparer.Create(Worksheet.Workbook.Culture, ignoreCase: true));
        for (var c = 0; c < names.Length; c++)
            columns.TryAdd(names[c], area.FirstColumn + c);
        for (var row = area.FirstRow + 1; row <= area.LastRow; row++)
        {
            for (var column = area.FirstColumn; column <= area.LastColumn; column++)
                Worksheet.CurrentValue(row, column);
        }

        var items = new List<T>(Math.Max(0, area.Rows - 1));
        for (var row = area.FirstRow + 1; row <= area.LastRow; row++)
            items.Add(read(new RangeRow(Worksheet, row, columns)));
        return items;
    }

    /// <summary>
    /// The range as CSV text, as Excel's Save As CSV writes it: each cell as it shows (with its number
    /// format), fields holding the delimiter, a quote or a line break in quotes, quotes doubled, and
    /// each row ending in CRLF. Whole columns such as <c>A:C</c> end at the last row used. Formulas out
    /// of date are calculated first.
    /// </summary>
    /// <param name="delimiter">What separates the fields; Excel uses <c>;</c> where the decimal separator is a comma.</param>
    /// <exception cref="ArgumentException">The delimiter is a quote or a line break.</exception>
    public string ToCsv(char delimiter = ',')
    {
        var writer = new StringWriter(CultureInfo.InvariantCulture);
        WriteCsv(writer, delimiter);
        return writer.ToString();
    }

    /// <summary>Writes the range as CSV text, for a file or a stream; see <see cref="ToCsv(char)"/>.</summary>
    /// <param name="writer">Where to write, in the encoding it has.</param>
    /// <param name="delimiter">What separates the fields.</param>
    /// <exception cref="ArgumentException">The delimiter is a quote or a line break.</exception>
    public void WriteCsv(TextWriter writer, char delimiter = ',')
    {
        ArgumentNullException.ThrowIfNull(writer);
        if (delimiter is '"' or '\r' or '\n')
            throw new ArgumentException("A quote or a line break cannot separate fields.", nameof(delimiter));
        var area = Exported();
        for (var row = area.FirstRow; row <= area.LastRow; row++)
        {
            for (var column = area.FirstColumn; column <= area.LastColumn; column++)
            {
                if (column > area.FirstColumn)
                    writer.Write(delimiter);
                Worksheet.CurrentValue(row, column);
                var text = new Cell(Worksheet, row, column).Text;
                if (text.AsSpan().IndexOfAny(['"', '\r', '\n', delimiter]) >= 0)
                    writer.Write("\"" + text.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"");
                else
                    writer.Write(text);
            }

            writer.Write("\r\n");
        }
    }

    // The range down to the rows and columns used, when it spans whole columns or rows.
    private Area Exported()
    {
        var area = Area;
        var used = Worksheet.Store.Bounds;
        if (area.FirstRow == 1 && area.LastRow == CellAddress.MaxRow)
            area = area with { LastRow = Math.Max(area.FirstRow, used?.LastRow ?? 1) };
        if (area.FirstColumn == 1 && area.LastColumn == CellAddress.MaxColumn)
            area = area with { LastColumn = Math.Max(area.FirstColumn, used?.LastColumn ?? 1) };
        return area;
    }

    // Column names from the first row, as tables name them; or Column1, Column2 and so on.
    private string[] ColumnNames(Area area, bool headers)
    {
        var names = new string[area.Columns];
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var c = 0; c < names.Length; c++)
        {
            var text = "";
            if (headers)
            {
                Worksheet.CurrentValue(area.FirstRow, area.FirstColumn + c);
                text = new Cell(Worksheet, area.FirstRow, area.FirstColumn + c).Text;
            }

            var name = text.Length == 0 ? "Column" + (c + 1).ToString(CultureInfo.InvariantCulture) : text;
            var unique = name;
            for (var n = 2; !used.Add(unique); n++)
                unique = name + n.ToString(CultureInfo.InvariantCulture);
            names[c] = unique;
        }

        return names;
    }
}
