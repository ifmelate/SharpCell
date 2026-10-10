using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using SharpCell.Evaluation;
using SharpCell.Parsing;

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
    private readonly CellMetadataWriter _metadata;
    private readonly StylesWriter _styles;
    private readonly SheetSource _source;

    // Some formula of the sheet changed: shared formulas are written in full, since the cell
    // holding a group's text may be the one that changed.
    private readonly bool _expandShared;
    private readonly IEnumerator<ModelCell> _model;
    private bool _hasModel;
    private bool _changed;
    private XNamespace _ns = XNamespace.None;
    private string _prefix = "";
    private CellAddress _previous;

    // Areas of kept dynamic arrays as the file states them: their spilled cells stay as saved.
    private readonly List<Area> _keptSpills = [];

    // The rectangle of the sheet's stored cells, or null for an empty sheet.
    private readonly Area? _used;

    private WorksheetWriter(Worksheet sheet, WritePlan plan, SavedCells saved, CellMetadataWriter metadata, StylesWriter styles)
    {
        _sheet = sheet;
        _plan = plan;
        _saved = saved;
        _metadata = metadata;
        _styles = styles;
        _source = plan.Source.Sheets[sheet];
        _expandShared = plan.FormulasChangedIn(sheet);
        _model = ModelCells(sheet).GetEnumerator();
        _hasModel = _model.MoveNext();
        foreach (var cell in ModelCells(sheet))
            _used = _used is { } area ? Area.Bounding(area, Area.Cell(cell.Row, cell.Column)) : Area.Cell(cell.Row, cell.Column);
    }

    /// <returns>The part written again, or null when no value in it changed.</returns>
    public static byte[]? Write(Package package, string part, Worksheet sheet, WritePlan plan, SavedCells saved,
        CellMetadataWriter metadata, StylesWriter styles)
    {
        var writer = new WorksheetWriter(sheet, plan, saved, metadata, styles);
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
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row" && reader.GetAttribute("r") is { } r)
                {
                    row = ParseRow(r);
                    NewRowsBefore(row, writer);
                    Row(reader, writer, row);
                }
                else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row")
                {
                    // Without its own number a row takes its first cell's, as the reader does (or
                    // follows the previous row): read it whole to find out before writing it.
                    var element = (XElement)XNode.ReadFrom(reader);
                    var first = element.Elements(_ns + "c").Select(c => (string?)c.Attribute("r")).FirstOrDefault(a => a is not null);
                    row = first is not null && CellAddress.TryParse(first, out var address) ? address.Row : row + 1;
                    NewRowsBefore(row, writer);
                    using var rowReader = element.CreateReader();
                    rowReader.MoveToContent();
                    Row(rowReader, writer, row);
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
                    if (source.Attribute("r") is { } r && CellAddress.TryParse(r.Value, out var address))
                    {
                        if (address.Row != row)
                            throw new InvalidDataException($"Cell {address} of sheet '{_sheet.Name}' lies in row {row}; the file cannot be written back.");
                        column = address.Column;
                    }
                    else
                    {
                        column++;
                    }

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

    // The sheet's stored cells merged with the positions that have only a style or a number format,
    // in row and column order.
    private static IEnumerable<ModelCell> ModelCells(Worksheet sheet)
    {
        var extras = new List<CellAddress>();
        foreach (var address in sheet.Styles.Keys)
        {
            if (sheet.Store.Get(address.Row, address.Column) is null)
                extras.Add(address);
        }

        foreach (var address in sheet.Formats.Keys)
        {
            if (sheet.Store.Get(address.Row, address.Column) is null && sheet.StyleAt(address.Row, address.Column) is null)
                extras.Add(address);
        }

        extras.Sort((a, b) => a.Row != b.Row ? a.Row.CompareTo(b.Row) : a.Column.CompareTo(b.Column));
        var next = 0;
        foreach (var cell in sheet.Store.Enumerate(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn))
        {
            while (next < extras.Count && (extras[next].Row < cell.Row || (extras[next].Row == cell.Row && extras[next].Column < cell.Column)))
            {
                yield return new ModelCell(extras[next].Row, extras[next].Column, null);
                next++;
            }

            yield return new ModelCell(cell.Row, cell.Column, cell.Data);
        }

        for (; next < extras.Count; next++)
            yield return new ModelCell(extras[next].Row, extras[next].Column, null);
    }

    private readonly record struct ModelCell(int Row, int Column, CellData? Data);

    // A cell the file does not have: a new constant or formula, a cell a spill now covers, or a
    // cell with only a style or a number format.
    private XElement? NewCell(ModelCell cell) => Styled(NewContent(cell), null, cell.Row, cell.Column);

    private XElement? NewContent(ModelCell cell)
    {
        if (cell.Data is not { } data)
            return null;
        var address = new CellAddress(cell.Row, cell.Column);
        if (data.FormulaText is not null)
            return WriteFormula(new XElement(_ns + "c", new XAttribute("r", address.ToString())), data, cell.Row, cell.Column);

        var value = CellEncoding.Storable(data.Value);
        if (value.Kind == CellValueKind.Empty
            || (data.SpillAnchor is not null && _keptSpills.Exists(area => area.Contains(cell.Row, cell.Column))))
            return null;
        CheckConstant(data, value, address);
        // A spilled cell holds its anchor's result, written like a formula's (text as t="str").
        return Rewrite(new XElement(_ns + "c", new XAttribute("r", address.ToString())), value, isFormula: data.SpillAnchor is not null, data);
    }

    // The source cell, the cell written again, or null when the cell goes away.
    private XElement? Cell(XElement source, int row, int column) =>
        Styled(Content(source, row, column), StyleIndex(source), row, column);

    private static int? StyleIndex(XElement cell) =>
        int.TryParse((string?)cell.Attribute("s"), NumberStyles.None, CultureInfo.InvariantCulture, out var index) ? index : null;

    // The cell with the style index its look and number format need now; a cell left with neither
    // content nor style goes away.
    private XElement? Styled(XElement? element, int? sourceStyle, int row, int column)
    {
        var target = TargetStyle(sourceStyle, row, column);
        if (target == sourceStyle)
            return element;
        _changed = true;
        if (element is null)
        {
            return target is { } index
                ? new XElement(_ns + "c", new XAttribute("r", new CellAddress(row, column).ToString()), new XAttribute("s", index))
                : null;
        }

        element.SetAttributeValue("s", target);
        return target is null && !element.HasElements ? null : element;
    }

    // The style index is kept while the cell looks as it did when read; otherwise the style is found
    // or made from the cell's old one. Cells without a style of their own show the workbook default.
    private int? TargetStyle(int? sourceStyle, int row, int column)
    {
        var address = new CellAddress(row, column);
        var style = _sheet.StyleAt(row, column) ?? _sheet.Workbook.DefaultStyle;
        var code = _sheet.FormatAt(row, column)?.FormatCode ?? "General";
        var loadedStyle = _source.Styles.TryGetValue(address, out var own) ? own : _plan.Source.DefaultStyle;
        var loadedCode = _source.Formats.TryGetValue(address, out var format) ? format : "General";
        if (style == loadedStyle && code == loadedCode)
            return sourceStyle;
        var index = _styles.XfFor(sourceStyle ?? 0, style, code);
        return index == 0 && sourceStyle is null ? null : index;
    }

    // The source cell with the model's content, or null when it has none.
    private XElement? Content(XElement source, int row, int column)
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

        // A kept spill keeps its saved cells, but not over a value the user typed there.
        if ((data is null || data.SpillAnchor is not null) && _keptSpills.Exists(area => area.Contains(row, column)))
            return source;

        var saved = _saved.Read(source, _sheet, row, column);
        var sourceFormula = source.Element(_ns + "f");
        if (data?.FormulaText is not null
            && (!WritePlan.IsLoadedFormula(data, _source, address) || (_expandShared && (string?)sourceFormula?.Attribute("t") == "shared")))
            return WriteFormula(source, data, row, column);

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

        // A formula the cell no longer has goes, even when its value stays. Excel saves the cells of a
        // volatile spill with an empty formula (<f ca="1"/>): those belong to the spill.
        var formulaRemoved = sourceFormula is not null && data?.SpillAnchor is null;

        // A #SPILL! or #CALC! value stays as the file has it: it cannot be written as a constant.
        if (unchanged && !formulaRemoved && (!_plan.RewriteAllValues || IsRichError(value)))
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
        {
            source.Elements(_ns + "f").Remove();
            source.Attribute("cm")?.Remove();
        }

        return Rewrite(source, value, isFormula: data?.SpillAnchor is not null, data);
    }

    // The cell with the model's formula in file syntax and its result.
    private XElement WriteFormula(XElement cell, CellData data, int row, int column)
    {
        cell.Elements(_ns + "f").Remove();
        cell.Attribute("cm")?.Remove();
        var origin = new CellAddress(row, column);
        var text = FormulaPrinter.PrintForFile(data.Formula!, origin, at => _sheet.Workbook.TableAt(_sheet, at.Row, at.Column)?.Name);
        var formula = new XElement(_ns + "f", text);
        if (data.FixedArray is { } fixedArea)
        {
            formula.Add(new XAttribute("t", "array"), new XAttribute("ref", Format(fixedArea)));
        }
        else if (!data.IsLegacy
                 && (data.SpillArea is not null || data.SpillWatch is not null || !FormulaStorage.IsPlain(data.Formula!, _sheet.Workbook, _sheet)))
        {
            // A dynamic array formula: its ref is the area it spills over, or just its cell.
            formula.Add(new XAttribute("t", "array"), new XAttribute("ref", Format(data.SpillArea ?? Area.Cell(row, column))));
            cell.SetAttributeValue("cm", _metadata.DynamicArrayCm());
        }

        cell.AddFirst(formula);
        return Rewrite(cell, CellEncoding.Storable(data.Value), isFormula: true, data);
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
            cell.SetAttributeValue("vm", _metadata.VmOf(error, wanted, wanted is null && isFormula && ReadsError(data!, error)));
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
