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
    private readonly IEnumerator<StoredCell> _model;
    private bool _hasModel;
    private bool _changed;
    private XNamespace _ns = XNamespace.None;
    private string _prefix = "";
    private CellAddress _previous;

    private WorksheetWriter(Worksheet sheet, WritePlan plan, SavedCells saved)
    {
        _sheet = sheet;
        _plan = plan;
        _saved = saved;
        _model = sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn).GetEnumerator();
        _hasModel = _model.MoveNext();
    }

    /// <returns>The part written again, or null when no value in it changed.</returns>
    public static byte[]? Write(Package package, string part, Worksheet sheet, WritePlan plan, SavedCells saved)
    {
        var writer = new WorksheetWriter(sheet, plan, saved);
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
        if (value.Kind == CellValueKind.Empty)
            return null;
        var address = new CellAddress(cell.Row, cell.Column);
        CheckConstant(cell.Data, value, address);
        return Rewrite(new XElement(_ns + "c", new XAttribute("r", address.ToString())), value, isFormula: false);
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

        var saved = _saved.Read(source, _sheet, row, column);
        if (data?.FormulaText is not null)
        {
            if (_plan.KeptCells.Contains(new CellKey(_sheet, row, column)))
                return source;
            var result = CellEncoding.Storable(data.Value);
            return !_plan.RewriteAllValues && result.Equals(saved.Value)
                ? source
                : Rewrite(source, result, isFormula: true);
        }

        var value = data is null ? CellValue.Empty : CellEncoding.Storable(data.Value);

        // A data table: Excel calculates it and the reader takes its value as a constant.
        if (saved.HasFormula)
        {
            return value.Equals(saved.Value)
                ? source
                : throw new NotSupportedException($"{_sheet.Name}!{address} is calculated by an Excel data table; its value cannot be changed.");
        }

        if (!_plan.RewriteAllValues && value.Equals(saved.Value))
            return source;
        if (value.Kind == CellValueKind.Empty)
        {
            _changed = true;
            return source.Attribute("s") is null
                ? null
                : new XElement(source.Name, source.Attributes().Where(a => a.Name.LocalName is not ("t" or "vm" or "cm")));
        }

        CheckConstant(data, value, address);
        return Rewrite(source, value, isFormula: false);
    }

    // A file can hold #SPILL! and #CALC! only as results; a typed one has nothing to stand for.
    private void CheckConstant(CellData? data, CellValue value, CellAddress address)
    {
        if (data?.SpillAnchor is null && value.Kind == CellValueKind.Error && CellEncoding.IsRichError(value.AsError()))
            throw new NotSupportedException($"{_sheet.Name}!{address} holds {value} as a value; files store #SPILL! and #CALC! only as formula results.");
    }

    private XElement Rewrite(XElement cell, CellValue value, bool isFormula)
    {
        _changed = true;
        cell.Attribute("t")?.Remove();
        cell.Attribute("vm")?.Remove();
        cell.Elements(_ns + "v").Remove();
        cell.Elements(_ns + "is").Remove();
        if (CellEncoding.TypeOf(value, isFormula) is { } type)
            cell.SetAttributeValue("t", type);
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
