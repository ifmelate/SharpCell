using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using SharpCell.Evaluation;

namespace SharpCell.Xlsx;

/// <summary>
/// Writes a worksheet part again with the workbook's values. Everything outside sheetData is copied
/// as it is, and so is every cell whose value did not change. The source's rows and cells are merged
/// with the sheet's stored cells, both in row and column order. Rows and cells get an explicit
/// position (r), so a new cell cannot shift neighbours that relied on following the previous one.
/// </summary>
internal sealed class WorksheetWriter
{
    private readonly Worksheet _sheet;
    private readonly WritePlan _plan;
    private readonly SavedCells _saved;
    private readonly RichErrors _richErrors;
    private readonly IEnumerator<StoredCell> _model;
    private bool _hasModel;
    private bool _changed;
    private XNamespace _ns = XNamespace.None;
    private string _prefix = "";
    private CellAddress _previous;

    // Areas of kept dynamic arrays as the file states them: their spilled cells stay as saved.
    private readonly List<Area> _keptSpills = [];

    // The rectangle of the sheet's stored cells, or null for an empty sheet.
    private readonly Area? _used;

    private WorksheetWriter(Worksheet sheet, WritePlan plan, SavedCells saved, RichErrors richErrors)
    {
        _sheet = sheet;
        _plan = plan;
        _saved = saved;
        _richErrors = richErrors;
        _model = sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn).GetEnumerator();
        _hasModel = _model.MoveNext();
        foreach (var cell in sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
            _used = _used is { } area ? Area.Bounding(area, Area.Cell(cell.Row, cell.Column)) : Area.Cell(cell.Row, cell.Column);
    }

    /// <returns>The part written again, or null when no value in it changed.</returns>
    public static byte[]? Write(Package package, string part, Worksheet sheet, WritePlan plan, SavedCells saved, RichErrors richErrors)
    {
        var writer = new WorksheetWriter(sheet, plan, saved, richErrors);
        var output = new MemoryStream();
        using (var reader = package.OpenXml(part))
        using (var xml = XmlWriter.Create(output, XmlParts.WriterSettings))
            writer.Copy(reader, xml);
        return writer._changed ? output.ToArray() : null;
    }

    private void Copy(XmlReader reader, XmlWriter writer)
    {
        reader.Read();
        while (!reader.EOF)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.Depth == 0)
            {
                var empty = reader.IsEmptyElement;
                XmlParts.CopyStartTag(reader, writer);
                if (empty)
                    writer.WriteEndElement();
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1 && reader.LocalName == "dimension")
            {
                Dimension((XElement)XNode.ReadFrom(reader)).WriteTo(writer);
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1 && reader.LocalName == "sheetData")
            {
                SheetData(reader, writer);
            }
            else
            {
                XmlParts.CopyNode(reader, writer);
            }
        }
    }

    // Widened to the cells the sheet now has; never narrowed (a larger one is also valid).
    private XElement Dimension(XElement dimension)
    {
        if (dimension.Attribute("ref") is { } reference && _used is { } used && TryParseArea(reference.Value, out var area))
        {
            var widened = Area.Bounding(area, used);
            if (widened != area)
                reference.Value = Format(widened);
        }

        return dimension;
    }

    private static bool TryParseArea(string text, out Area area)
    {
        var colon = text.IndexOf(':');
        var first = colon < 0 ? text.AsSpan() : text.AsSpan(0, colon);
        var last = colon < 0 ? text.AsSpan() : text.AsSpan(colon + 1);
        if (CellAddress.TryParse(first, out var a) && CellAddress.TryParse(last, out var b))
        {
            area = new Area(Math.Min(a.Row, b.Row), Math.Min(a.Column, b.Column), Math.Max(a.Row, b.Row), Math.Max(a.Column, b.Column));
            return true;
        }

        area = default;
        return false;
    }

    private static string Format(Area area) => area.IsSingleCell
        ? new CellAddress(area.FirstRow, area.FirstColumn).ToString()
        : $"{new CellAddress(area.FirstRow, area.FirstColumn)}:{new CellAddress(area.LastRow, area.LastColumn)}";

    private void SheetData(XmlReader reader, XmlWriter writer)
    {
        _ns = reader.NamespaceURI;
        _prefix = reader.Prefix;
        var depth = reader.Depth;
        var empty = reader.IsEmptyElement;
        XmlParts.CopyStartTag(reader, writer);
        var row = 0;
        if (!empty)
        {
            while (!(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row")
                {
                    row = reader.GetAttribute("r") is { } r ? ParseRow(r) : row + 1;
                    NewRowsBefore(row, writer);
                    Row(reader, writer, row);
                }
                else
                {
                    XmlParts.CopyNode(reader, writer);
                }
            }

            reader.Read();
        }

        NewRowsBefore(int.MaxValue, writer);
        writer.WriteFullEndElement();
    }

    private void Row(XmlReader reader, XmlWriter writer, int row)
    {
        var depth = reader.Depth;
        var empty = reader.IsEmptyElement;
        var positioned = reader.GetAttribute("r") is not null;
        XmlParts.CopyStartTag(reader, writer);
        if (!positioned)
            writer.WriteAttributeString("r", row.ToString(CultureInfo.InvariantCulture));

        var column = 0;
        if (!empty)
        {
            while (!(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
                {
                    var source = (XElement)XNode.ReadFrom(reader);
                    column = source.Attribute("r") is { } r && CellAddress.TryParse(r.Value, out var address) ? address.Column : column + 1;
                    CheckOrder(row, column);
                    NewCellsBefore(row, column, writer);
                    Cell(source, row, column)?.WriteTo(writer);
                }
                else
                {
                    XmlParts.CopyNode(reader, writer);
                }
            }

            reader.Read();
        }

        NewCellsBefore(row, int.MaxValue, writer);
        writer.WriteFullEndElement();
    }

    // Merging needs the file's cells in order, as Excel writes them.
    private void CheckOrder(int row, int column)
    {
        var current = new CellAddress(row, column);
        if (_previous != default && (row < _previous.Row || (row == _previous.Row && column <= _previous.Column)))
            throw new InvalidDataException($"Cells of sheet '{_sheet.Name}' are not in order at {current}; the file cannot be written back.");
        _previous = current;
    }

    private void NewRowsBefore(int row, XmlWriter writer)
    {
        while (_hasModel && _model.Current.Row < row)
        {
            var current = _model.Current.Row;
            var cells = new List<XElement>();
            while (_hasModel && _model.Current.Row == current)
            {
                if (NewCell(_model.Current) is { } cell)
                    cells.Add(cell);
                Advance();
            }

            if (cells.Count == 0)
                continue;
            writer.WriteStartElement(_prefix, "row", _ns.NamespaceName);
            writer.WriteAttributeString("r", current.ToString(CultureInfo.InvariantCulture));
            foreach (var cell in cells)
                cell.WriteTo(writer);
            writer.WriteEndElement();
        }
    }

    private void NewCellsBefore(int row, int column, XmlWriter writer)
    {
        while (_hasModel && _model.Current.Row == row && _model.Current.Column < column)
        {
            NewCell(_model.Current)?.WriteTo(writer);
            Advance();
        }
    }

    private void Advance() => _hasModel = _model.MoveNext();

    // A stored cell the file does not have: a new constant, or a cell a spill now covers.
    private XElement? NewCell(StoredCell cell)
    {
        var value = CellEncoding.Storable(cell.Data.Value);
        if (value.Kind == CellValueKind.Empty || _keptSpills.Exists(area => area.Contains(cell.Row, cell.Column)))
            return null;
        var address = new CellAddress(cell.Row, cell.Column);
        CheckConstant(cell.Data, value, address);
        return Rewrite(new XElement(_ns + "c", new XAttribute("r", address.ToString())), value, isFormula: false, cell.Data);
    }

    // The source cell, the cell written again, or null when the cell goes away.
    private XElement? Cell(XElement source, int row, int column)
    {
        var address = new CellAddress(row, column);
        if (source.Attribute("r") is null)
            source.ReplaceAttributes(new XAttribute("r", address.ToString()), source.Attributes().ToList());

        CellData? data = null;
        if (_hasModel && _model.Current.Row == row && _model.Current.Column == column)
        {
            data = _model.Current.Data;
            Advance();
        }

        if (_keptSpills.Exists(area => area.Contains(row, column)))
            return source;

        var saved = _saved.Read(source, _sheet, row, column);
        if (data?.FormulaText is not null)
        {
            if (_plan.KeptCells.Contains(new CellKey(_sheet, row, column)))
            {
                if (source.Element(_ns + "f") is { } kept && (string?)kept.Attribute("t") == "array"
                    && (string?)kept.Attribute("ref") is { } reference && TryParseArea(reference, out var spill))
                    _keptSpills.Add(spill);
                return source;
            }
            var moved = UpdateSpillRef(source, data, row, column);
            var result = CellEncoding.Storable(data.Value);
            return !moved && !_plan.RewriteAllValues && result.Equals(saved.Value)
                ? source
                : Rewrite(source, result, isFormula: true, data);
        }

        var value = data is null ? CellValue.Empty : CellEncoding.Storable(data.Value);
        var unchanged = value.Equals(saved.Value);

        // A data table: Excel calculates it and the reader takes its value as a constant.
        if (saved.IsDataTable)
        {
            return unchanged
                ? source
                : throw new NotSupportedException($"{_sheet.Name}!{address} is calculated by an Excel data table; its value cannot be changed.");
        }

        // A #SPILL! or #CALC! value stays as the file has it: it cannot be written as a constant.
        if (unchanged && (!_plan.RewriteAllValues || IsRichError(value)))
            return source;
        if (value.Kind == CellValueKind.Empty)
        {
            _changed = true;
            return source.Attribute("s") is null
                ? null
                : new XElement(source.Name, source.Attributes().Where(a => a.Name.LocalName is not ("t" or "vm" or "cm")));
        }

        CheckConstant(data, value, address);

        // Excel saves the cells of a volatile spill with an empty formula (<f ca="1"/>): a value typed
        // over one is the user's, without it.
        if (data?.SpillAnchor is null)
            source.Elements(_ns + "f").Remove();
        return Rewrite(source, value, isFormula: false, data);
    }

    // A dynamic array formula's ref is the area its result spills over, or just its cell.
    private bool UpdateSpillRef(XElement cell, CellData data, int row, int column)
    {
        if (data.FixedArray is not null || data.IsLegacy || cell.Element(_ns + "f") is not { } formula || (string?)formula.Attribute("t") != "array")
            return false;
        var text = Format(data.SpillArea ?? Area.Cell(row, column));
        if ((string?)formula.Attribute("ref") == text)
            return false;
        formula.SetAttributeValue("ref", text);
        return true;
    }

    // Whether the formula read a cell holding the error: then it passes the error on rather than
    // causing it, which Excel records as a propagated error.
    private static bool ReadsError(CellData data, ErrorKind error)
    {
        foreach (var (sheet, area) in data.Registered?.Areas ?? [])
        {
            foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
            {
                if (cell.Data.Value.Kind == CellValueKind.Error && cell.Data.Value.AsError() == error)
                    return true;
            }
        }

        return false;
    }

    private static bool IsRichError(CellValue value) => value.Kind == CellValueKind.Error && CellEncoding.IsRichError(value.AsError());

    // A file can hold #SPILL! and #CALC! only as results; a typed one has nothing to stand for.
    private void CheckConstant(CellData? data, CellValue value, CellAddress address)
    {
        if (data?.SpillAnchor is null && value.Kind == CellValueKind.Error && CellEncoding.IsRichError(value.AsError()))
            throw new NotSupportedException($"{_sheet.Name}!{address} holds {value} as a value; files store #SPILL! and #CALC! only as formula results.");
    }

    private XElement Rewrite(XElement cell, CellValue value, bool isFormula, CellData? data)
    {
        _changed = true;
        cell.Attribute("t")?.Remove();
        cell.Attribute("vm")?.Remove();
        cell.Elements(_ns + "v").Remove();
        cell.Elements(_ns + "is").Remove();
        if (CellEncoding.TypeOf(value, isFormula) is { } type)
            cell.SetAttributeValue("t", type);
        if (value.Kind == CellValueKind.Error && CellEncoding.IsRichError(value.AsError()))
        {
            // A #SPILL! anchor that cells block remembers the area it wanted (SpillWatch); at the
            // sheet's edge or in a table it has none.
            var error = value.AsError();
            var wanted = isFormula && error == ErrorKind.Spill ? data?.SpillWatch : null;
            cell.SetAttributeValue("vm", _richErrors.VmOf(error, wanted, wanted is null && isFormula && ReadsError(data!, error)));
        }

        if (CellEncoding.ValueElement(_ns, value, isFormula) is { } element)
        {
            if (cell.Element(_ns + "f") is { } formula)
                formula.AddAfterSelf(element);
            else
                cell.AddFirst(element);
        }

        return cell;
    }

    private int ParseRow(string text) =>
        int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var row) && row is >= 1 and <= CellAddress.MaxRow
            ? row
            : throw new InvalidDataException($"Row number '{text}' in sheet '{_sheet.Name}' is not valid.");
}
