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

    // What changed apart from cells, against the sheet as read: those parts are written again.
    private readonly SheetLook _look;
    private readonly bool _columnsChanged;
    private readonly bool _rowsChanged;
    private readonly bool _mergesChanged;
    private readonly bool _viewChanged;
    private readonly bool _defaultsChanged;
    private readonly bool _filterChanged;

    // The data validation rules are written from the model when they changed, and always on a sheet
    // whose rows or columns moved: the model's rules have followed the edits, so the file's element is
    // never moved (rules the reader skipped as damaged go rather than stay where they no longer belong).
    private readonly bool _validationsChanged;

    // Elements to add where the schema puts them when the part has none.
    private readonly List<string> _pending = [];

    // Rows whose height or hidden flag changed, in order; those the file has no element for get one.
    private readonly List<int> _changedRows = [];
    private int _nextChangedRow;

    // Shared groups whose first cell (holding their text) was changed or removed.
    private readonly HashSet<string> _brokenGroups = new(StringComparer.Ordinal);

    // Areas of kept dynamic arrays as the file states them: their spilled cells stay as saved.
    private readonly List<Area> _keptSpills = [];

    // The rectangle of the sheet's stored cells, or null for an empty sheet.
    private readonly Area? _used;

    // Relationship ids of table parts added for new tables, and of those of removed tables.
    private readonly IReadOnlyList<string> _newTables;
    private readonly IReadOnlySet<string> _removedTables;

    // Rows and columns inserted or deleted: in this sheet, the file's positions move; in any
    // sheet, formulas the model does not hold (conditional formats) follow the references.
    private readonly SheetEditMap _map;
    private readonly bool _shifted;

    // Header cells of new tables, written with their column's name.
    private readonly IReadOnlyDictionary<CellAddress, string> _headers;

    // The links to write when they changed; null when the file's element stays (moved by the edits).
    private readonly SheetHyperlinks? _hyperlinks;

    private WorksheetWriter(Worksheet sheet, SheetSource source, WritePlan plan, SavedCells saved, CellMetadataWriter metadata,
        StylesWriter styles, IReadOnlyList<string> newTables, IReadOnlySet<string> removedTables, IReadOnlyDictionary<CellAddress, string> headers,
        SheetHyperlinks? hyperlinks)
    {
        _map = plan.Edits;
        _shifted = source.Part is not null && _map.HasEdits(sheet);
        _removedTables = removedTables;
        _sheet = sheet;
        _plan = plan;
        _saved = saved;
        _metadata = metadata;
        _styles = styles;
        _source = source;
        _newTables = newTables;
        _headers = headers;
        _hyperlinks = hyperlinks;
        _expandShared = plan.FormulasChangedIn(sheet);
        if (newTables.Count > 0 || removedTables.Count > 0)
            _pending.Add("tableParts");
        _model = ModelCells(sheet, headers.Keys).GetEnumerator();
        _hasModel = _model.MoveNext();
        foreach (var cell in ModelCells(sheet, headers.Keys))
            _used = _used is { } area ? Area.Bounding(area, Area.Cell(cell.Row, cell.Column)) : Area.Cell(cell.Row, cell.Column);

        _look = _source.Look.Mapped(_map, sheet);
        _columnsChanged = _look.ColumnsChanged(sheet) || (_shifted && _map.HasColumnEdits(sheet));
        _rowsChanged = _look.RowsChanged(sheet);
        _mergesChanged = _look.MergesChanged(sheet) || _shifted;
        _viewChanged = _look.ViewChanged(sheet);
        _defaultsChanged = _look.DefaultsChanged(sheet);
        _filterChanged = _look.FilterChanged(sheet);
        _validationsChanged = _shifted || _look.ValidationsChanged(sheet);
        if (_filterChanged && sheet.FilterMode)
            _pending.Add("sheetPr");
        if (_viewChanged)
            _pending.Add("sheetViews");
        if (_defaultsChanged)
            _pending.Add("sheetFormatPr");
        if (_columnsChanged)
            _pending.Add("cols");
        if (_mergesChanged)
            _pending.Add("mergeCells");
        if (_validationsChanged)
            _pending.Add("dataValidations");
        if (hyperlinks is not null)
            _pending.Add("hyperlinks");
        if (_rowsChanged)
        {
            var rows = new SortedSet<int>(_look.RowHeights.Keys);
            rows.UnionWith(_look.HiddenRows);
            rows.UnionWith(sheet.RowHeights.Keys);
            rows.UnionWith(sheet.HiddenRows);
            _changedRows.AddRange(rows.Where(row => _look.RowChanged(sheet, row)));
        }
    }

    /// <returns>The part written again, or null when no value in it changed.</returns>
    /// <param name="reader">The worksheet part as the file has it, or an empty one for a new sheet.</param>
    /// <param name="sheet">The sheet.</param>
    /// <param name="source">The sheet as read.</param>
    /// <param name="plan">The save.</param>
    /// <param name="saved">What the file's cells point into.</param>
    /// <param name="metadata">Cell metadata to find or add entries in.</param>
    /// <param name="styles">Cell styles to find or add entries in.</param>
    /// <param name="newTables">Relationship ids of table parts to list.</param>
    /// <param name="headers">Header cells of new tables, with their column names.</param>
    /// <param name="removedTables">Relationship ids of table parts to leave out.</param>
    /// <param name="hyperlinks">The links to write, or null to keep the file's.</param>
    public static byte[]? Write(XmlReader reader, Worksheet sheet, SheetSource source, WritePlan plan, SavedCells saved,
        CellMetadataWriter metadata, StylesWriter styles, IReadOnlyList<string> newTables, IReadOnlySet<string> removedTables,
        IReadOnlyDictionary<CellAddress, string> headers, SheetHyperlinks? hyperlinks)
    {
        var writer = new WorksheetWriter(sheet, source, plan, saved, metadata, styles, newTables, removedTables, headers, hyperlinks);
        var output = new MemoryStream();
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
                _ns = reader.NamespaceURI;
                _prefix = reader.Prefix;
                var empty = reader.IsEmptyElement;
                XmlParts.CopyStartTag(reader, writer);
                if (empty)
                {
                    WritePending(int.MaxValue, writer);
                    writer.WriteEndElement();
                }
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1)
            {
                var name = reader.LocalName;
                var rank = reader.NamespaceURI == _ns.NamespaceName ? WorksheetParts.Rank(name) : -1;
                if (rank >= 0)
                    WritePending(rank, writer);
                if (rank < 0)
                {
                    XmlParts.CopyNode(reader, writer);
                    continue;
                }

                switch (name)
                {
                    case "dimension":
                        Dimension((XElement)XNode.ReadFrom(reader)).WriteTo(writer);
                        break;
                    case "sheetData":
                        SheetData(reader, writer);
                        break;
                    case "sheetViews" when !_viewChanged && _shifted:
                    case "autoFilter" when _shifted:
                    case "hyperlinks" when _shifted && _hyperlinks is null:
                    case "rowBreaks" when _shifted:
                    case "colBreaks" when _shifted:
                    case "ignoredErrors" when _shifted:
                    case "protectedRanges" when _shifted:
                    case "conditionalFormatting" when _map.Any:
                    case "extLst" when _map.Any:
                        Shifted((XElement)XNode.ReadFrom(reader))?.WriteTo(writer);
                        break;
                    case "tableParts" when _newTables.Count > 0 || _removedTables.Count > 0:
                    case "sheetPr" when _filterChanged:
                    case "sheetViews" when _viewChanged:
                    case "sheetFormatPr" when _defaultsChanged:
                    case "cols" when _columnsChanged:
                    case "mergeCells" when _mergesChanged:
                    case "dataValidations" when _validationsChanged:
                    case "hyperlinks" when _hyperlinks is not null:
                        _pending.Remove(name);
                        _changed = true;
                        Build(name, (XElement)XNode.ReadFrom(reader))?.WriteTo(writer);
                        break;
                    default:
                        XmlParts.CopyNode(reader, writer);
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == 0)
            {
                WritePending(int.MaxValue, writer);
                XmlParts.CopyNode(reader, writer);
            }
            else
            {
                XmlParts.CopyNode(reader, writer);
            }
        }
    }

    // Adds the elements the part lacks that come before an element of this rank.
    private void WritePending(int rank, XmlWriter writer)
    {
        foreach (var name in _pending.Where(name => WorksheetParts.Rank(name) < rank).OrderBy(WorksheetParts.Rank).ToList())
        {
            _pending.Remove(name);
            if (Build(name, null) is { } element)
            {
                _changed = true;
                element.WriteTo(writer);
            }
        }
    }

    // The element as the sheet now needs it, from the file's (or none); null when it has nothing to say.
    private XElement? Build(string name, XElement? source) => name switch
    {
        "sheetPr" => SheetProperties(source),
        "sheetViews" when _shifted => Shifted(SheetViews(source)),
        "sheetViews" => SheetViews(source),
        "sheetFormatPr" => SheetFormat(source),
        "cols" => Columns(source),
        "tableParts" => TableParts(source),
        "dataValidations" => ValidationElements.Element(_ns, _sheet, source),
        "hyperlinks" => Hyperlinks(),
        _ => MergeCells(),
    };

    // An element of the file holding positions or formulas, with them moved by the edits; null when
    // nothing of it is left.
    private XElement? Shifted(XElement? element)
    {
        if (element is null)
            return null;
        var before = new XElement(element);
        var result = element.Name.LocalName switch
        {
            "autoFilter" => AutoFilter(element),
            "sheetViews" => Selections(element),
            "conditionalFormatting" => Ranged(element),
            "hyperlinks" => List(element, "hyperlink", e => MapAttribute(e, "ref")),
            "ignoredErrors" => List(element, "ignoredError", e => MapAttribute(e, "sqref")),
            "protectedRanges" => List(element, "protectedRange", e => MapAttribute(e, "sqref")),
            "rowBreaks" => Breaks(element, rows: true),
            "colBreaks" => Breaks(element, rows: false),
            "extLst" => Extensions(element),
            _ => element,
        };

        if (result is null || !XNode.DeepEquals(before, result))
            _changed = true;
        return result;
    }

    private XElement? MapAttribute(XElement element, string name)
    {
        if (!_shifted || (string?)element.Attribute(name) is not { } text)
            return element;
        var mapped = _map.MapSqref(_sheet, text);
        if (mapped is null)
            return null;
        element.SetAttributeValue(name, mapped);
        return element;
    }

    // An element with a sqref and formulas relative to its first cell (a conditional format): the
    // formulas follow the references, then the sqref moves.
    private XElement? Ranged(XElement element)
    {
        var sqref = (string?)element.Attribute("sqref");
        if (SheetEditMap.FirstCell(sqref) is { } origin)
        {
            var anchor = _shifted ? _map.FirstSurvivingCell(_sheet, sqref!) : origin;
            foreach (var formula in element.Descendants().Where(d => d.Name.LocalName is "formula" or "formula1" or "formula2" && !d.HasElements))
                formula.Value = _map.MapFormula(formula.Value, _sheet, origin, anchor);
        }

        return MapAttribute(element, "sqref");
    }

    private static XElement? List(XElement element, string child, Func<XElement, XElement?> map)
    {
        foreach (var item in element.Elements(element.Name.Namespace + child).ToList())
        {
            if (map(item) is null)
                item.Remove();
        }

        var count = element.Elements(element.Name.Namespace + child).Count();
        if (count == 0)
            return null;
        if (element.Attribute("count") is not null)
            element.SetAttributeValue("count", count);
        return element;
    }

    private XElement? AutoFilter(XElement element)
    {
        if ((string?)element.Attribute("ref") is not { } text || !TryParseArea(text, out var area))
            return element;
        if (_map.Map(_sheet, area) is not { } moved)
            return null;
        element.SetAttributeValue("ref", Format(moved));

        // Filter columns count from the filter's first column.
        foreach (var column in element.Elements(_ns + "filterColumn").ToList())
        {
            if (!int.TryParse((string?)column.Attribute("colId"), NumberStyles.None, CultureInfo.InvariantCulture, out var id))
                continue;
            if (_map.MapColumn(_sheet, area.FirstColumn + id) is { } to)
                column.SetAttributeValue("colId", to - moved.FirstColumn);
            else
                column.Remove();
        }

        return element;
    }

    private XElement Selections(XElement element)
    {
        foreach (var selection in element.Descendants(_ns + "selection"))
        {
            if ((string?)selection.Attribute("activeCell") is { } active && CellAddress.TryParse(active, out var cell))
                selection.SetAttributeValue("activeCell", _map.Map(_sheet, cell)?.ToString());
            if ((string?)selection.Attribute("sqref") is { } sqref)
                selection.SetAttributeValue("sqref", _map.MapSqref(_sheet, sqref));
        }

        return element;
    }

    // A page break's id is the last row (or column) before it.
    private XElement? Breaks(XElement element, bool rows)
    {
        foreach (var brk in element.Elements(_ns + "brk").ToList())
        {
            if (!int.TryParse((string?)brk.Attribute("id"), NumberStyles.None, CultureInfo.InvariantCulture, out var id))
                continue;
            if ((rows ? _map.MapRow(_sheet, id) : _map.MapColumn(_sheet, id)) is { } to)
                brk.SetAttributeValue("id", to);
            else
                brk.Remove();
        }

        var breaks = element.Elements(_ns + "brk").ToList();
        if (breaks.Count == 0)
            return null;
        element.SetAttributeValue("count", breaks.Count);
        if (element.Attribute("manualBreakCount") is not null)
            element.SetAttributeValue("manualBreakCount", breaks.Count(b => (string?)b.Attribute("man") is "1" or "true"));
        return element;
    }

    // Newer conditional formats and validations keep their ranges in xm:sqref and their formulas in xm:f.
    private XElement? Extensions(XElement element)
    {
        foreach (var sqref in element.Descendants().Where(d => d.Name.LocalName == "sqref" && !d.HasElements).ToList())
        {
            var owner = sqref.Parent!;
            if (SheetEditMap.FirstCell(sqref.Value) is { } origin)
            {
                var anchor = _shifted ? _map.FirstSurvivingCell(_sheet, sqref.Value) : origin;
                foreach (var formula in owner.Descendants().Where(d => d.Name.LocalName == "f" && !d.HasElements))
                    formula.Value = _map.MapFormula(formula.Value, _sheet, origin, anchor);
            }

            if (!_shifted)
                continue;
            if (_map.MapSqref(_sheet, sqref.Value) is { } mapped)
                sqref.Value = mapped;
            else
                owner.Remove();
        }

        // Lists left empty go, up to the extension list itself: the schema wants at least one item.
        foreach (var list in element.Descendants().Where(d => d.Name.LocalName is "conditionalFormattings" or "dataValidations" or "sparklines").ToList())
        {
            var count = list.Elements().Count();
            if (count == 0)
                list.Remove();
            else if (list.Attribute("count") is not null)
                list.SetAttributeValue("count", count);
        }

        element.Descendants().Where(d => d.Name.LocalName == "sparklineGroup" && !d.Elements().Any(e => e.Name.LocalName == "sparklines")).ToList().ForEach(d => d.Remove());
        element.Descendants().Where(d => d.Name.LocalName == "sparklineGroups" && !d.HasElements).ToList().ForEach(d => d.Remove());
        element.Elements().Where(e => e.Name.LocalName == "ext" && !e.HasElements).ToList().ForEach(e => e.Remove());
        return element.HasElements ? element : null;
    }

    private XElement? TableParts(XElement? element)
    {
        XNamespace relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        element ??= new XElement(_ns + "tableParts", new XAttribute(XNamespace.Xmlns + "r", relationships));
        var existing = element.Elements(_ns + "tablePart").SelectMany(t => t.Attributes()).FirstOrDefault(a => a.Name.LocalName == "id" && a.Name.Namespace != XNamespace.None);
        var r = existing?.Name.Namespace ?? relationships;
        element.Elements(_ns + "tablePart").Where(t => t.Attributes().Any(a => a.Name.LocalName == "id" && _removedTables.Contains(a.Value))).Remove();
        foreach (var id in _newTables)
            element.Add(new XElement(_ns + "tablePart", new XAttribute(r + "id", id)));
        var count = element.Elements(_ns + "tablePart").Count();
        element.SetAttributeValue("count", count);
        return count == 0 ? null : element;
    }

    private XElement SheetProperties(XElement? element)
    {
        element ??= new XElement(_ns + "sheetPr");
        element.SetAttributeValue("filterMode", _sheet.FilterMode ? "1" : null);
        return element;
    }

    private XElement SheetViews(XElement? element)
    {
        element ??= new XElement(_ns + "sheetViews");
        var view = element.Element(_ns + "sheetView");
        if (view is null)
        {
            view = new XElement(_ns + "sheetView", new XAttribute("workbookViewId", 0));
            element.AddFirst(view);
        }

        view.SetAttributeValue("showGridLines", _sheet.ShowGridlines ? null : "0");
        if (!_look.PanesChanged(_sheet))
            return element;

        // Frozen panes replace any pane and the selections that belonged to its parts.
        view.Elements(_ns + "pane").Remove();
        view.Elements(_ns + "selection").Where(selection => selection.Attribute("pane") is not null).Remove();
        int rows = _sheet.FrozenRows, columns = _sheet.FrozenColumns;
        if (rows == 0 && columns == 0)
            return element;
        var active = rows > 0 && columns > 0 ? "bottomRight" : rows > 0 ? "bottomLeft" : "topRight";
        var pane = new XElement(_ns + "pane",
            columns > 0 ? new XAttribute("xSplit", columns) : null,
            rows > 0 ? new XAttribute("ySplit", rows) : null,
            new XAttribute("topLeftCell", new CellAddress(rows + 1, columns + 1).ToString()),
            new XAttribute("activePane", active),
            new XAttribute("state", "frozen"));
        view.AddFirst(pane);
        pane.AddAfterSelf(new XElement(_ns + "selection", new XAttribute("pane", active)));
        return element;
    }

    private XElement SheetFormat(XElement? element)
    {
        element ??= new XElement(_ns + "sheetFormatPr");
        element.SetAttributeValue("defaultColWidth", _sheet.DefaultColumnWidth is { } width ? WorksheetParts.Number(width) : null);
        if (_sheet.DefaultRowHeight is { } height)
        {
            element.SetAttributeValue("defaultRowHeight", WorksheetParts.Number(height));
            element.SetAttributeValue("customHeight", "1");
        }
        else if (element.Attribute("defaultRowHeight") is null)
        {
            element.SetAttributeValue("defaultRowHeight", "15");   // required by the schema; Excel's for Calibri 11
        }

        return element;
    }

    // The file's column ranges cut into single columns, the widths and hidden flags that changed
    // applied, and joined again where neighbours agree; other attributes (styles) are kept.
    private XElement? Columns(XElement? element)
    {
        var columns = new SortedDictionary<int, List<(string Name, string Value)>>();
        foreach (var col in element?.Elements(_ns + "col") ?? [])
        {
            if (!int.TryParse((string?)col.Attribute("min"), NumberStyles.None, CultureInfo.InvariantCulture, out var min)
                || !int.TryParse((string?)col.Attribute("max"), NumberStyles.None, CultureInfo.InvariantCulture, out var max)
                || min < 1 || max > CellAddress.MaxColumn || min > max)
                continue;
            var attributes = col.Attributes().Where(a => a.Name.LocalName is not ("min" or "max")).Select(a => (a.Name.LocalName, a.Value)).ToList();
            for (var column = min; column <= max; column++)
            {
                if ((_shifted ? _map.MapColumn(_sheet, column) : column) is { } to)
                    columns[to] = [.. attributes];
            }
        }

        var touched = new SortedSet<int>(_look.ColumnWidths.Keys);
        touched.UnionWith(_look.HiddenColumns);
        touched.UnionWith(_sheet.ColumnWidths.Keys);
        touched.UnionWith(_sheet.HiddenColumns);
        foreach (var column in touched)
        {
            var width = _sheet.ColumnWidth(column);
            var hidden = _sheet.IsColumnHidden(column);
            if (!columns.TryGetValue(column, out var attributes))
                columns[column] = attributes = [];
            if (width != _look.LoadedWidth(column))
            {
                Set(attributes, "width", width is { } w ? WorksheetParts.Number(w) : null);
                Set(attributes, "customWidth", width is null ? null : "1");
            }

            if (hidden != _look.HiddenColumns.Contains(column))
                Set(attributes, "hidden", hidden ? "1" : null);
            if (attributes.Count == 0)
                columns.Remove(column);
        }

        if (columns.Count == 0)
            return null;
        var result = new XElement(_ns + "cols");
        int? start = null, previous = null;
        List<(string, string)>? run = null;
        void Flush()
        {
            if (start is { } first)
            {
                result.Add(new XElement(_ns + "col", new XAttribute("min", first), new XAttribute("max", previous!.Value),
                    run!.Select(a => new XAttribute(a.Item1, a.Item2))));
            }
        }

        foreach (var (column, attributes) in columns)
        {
            if (start is null || previous != column - 1 || !run!.SequenceEqual(attributes))
            {
                Flush();
                start = column;
                run = attributes;
            }

            previous = column;
        }

        Flush();
        return result;
    }

    private static void Set(List<(string Name, string Value)> attributes, string name, string? value)
    {
        var index = attributes.FindIndex(a => a.Name == name);
        if (value is null)
        {
            if (index >= 0)
                attributes.RemoveAt(index);
        }
        else if (index >= 0)
        {
            attributes[index] = (name, value);
        }
        else
        {
            attributes.Add((name, value));
        }
    }

    // The kept elements move with the rows and columns, as the file's element would.
    private XElement? Hyperlinks() => HyperlinkWriter.Element(_ns, _sheet, _hyperlinks!.Ids,
        _hyperlinks.Kept.Select(element => MapAttribute(new XElement(element), "ref")).OfType<XElement>());

    private XElement? MergeCells()
    {
        var merged = _sheet.MergedAreaList;
        return merged.Count == 0
            ? null
            : new XElement(_ns + "mergeCells", new XAttribute("count", merged.Count),
                merged.Select(area => new XElement(_ns + "mergeCell", new XAttribute("ref", Format(area)))));
    }

    private void WriteRowAttributes(XmlWriter writer, int row)
    {
        if (_sheet.RowHeight(row) is { } height)
        {
            writer.WriteAttributeString("ht", WorksheetParts.Number(height));
            writer.WriteAttributeString("customHeight", "1");
        }

        if (_sheet.IsRowHidden(row))
            writer.WriteAttributeString("hidden", "1");
    }

    // Widened to the cells the sheet now has; never narrowed (a larger one is also valid), except
    // after rows or columns moved.
    private XElement Dimension(XElement dimension)
    {
        if (dimension.Attribute("ref") is { } reference && _used is { } used && TryParseArea(reference.Value, out var area))
        {
            var widened = _shifted ? used : Area.Bounding(area, used);
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
                    if (MappedRow(row) is { } mapped)
                    {
                        NewRowsBefore(mapped, writer);
                        Row(reader, writer, mapped, row);
                    }
                    else
                    {
                        reader.Skip();
                    }
                }
                else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row")
                {
                    // Without its own number a row takes its first cell's, as the reader does (or
                    // follows the previous row): read it whole to find out before writing it.
                    var element = (XElement)XNode.ReadFrom(reader);
                    var first = element.Elements(_ns + "c").Select(c => (string?)c.Attribute("r")).FirstOrDefault(a => a is not null);
                    row = first is not null && CellAddress.TryParse(first, out var address) ? address.Row : row + 1;
                    if (MappedRow(row) is not { } mapped)
                        continue;
                    NewRowsBefore(mapped, writer);
                    using var rowReader = element.CreateReader();
                    rowReader.MoveToContent();
                    Row(rowReader, writer, mapped, row);
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

    // A row of the file where the edits put it; null when it was deleted.
    private int? MappedRow(int row)
    {
        if (!_shifted)
            return row;
        var mapped = _map.MapRow(_sheet, row);
        if (mapped is null)
            _changed = true;
        return mapped;
    }

    /// <param name="reader">On the row element.</param>
    /// <param name="writer">Where the row goes.</param>
    /// <param name="row">The row's number in the sheet now.</param>
    /// <param name="sourceRow">Its number in the file.</param>
    private void Row(XmlReader reader, XmlWriter writer, int row, int sourceRow)
    {
        var depth = reader.Depth;
        var empty = reader.IsEmptyElement;
        var positioned = reader.GetAttribute("r") is not null;
        while (_nextChangedRow < _changedRows.Count && _changedRows[_nextChangedRow] <= row)
            _nextChangedRow++;
        var sized = _rowsChanged && _look.RowChanged(_sheet, row);
        if (sized || _shifted)
        {
            // The row's number, height and hidden flag as the sheet has them now; its other attributes
            // stay, except spans, which moved columns would make wrong.
            _changed = true;
            writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
            writer.WriteAttributeString("r", row.ToString(CultureInfo.InvariantCulture));
            if (reader.MoveToFirstAttribute())
            {
                do
                {
                    var skip = reader.NamespaceURI.Length == 0
                        && (reader.LocalName == "r" || (_shifted && reader.LocalName == "spans") || (sized && reader.LocalName is "ht" or "customHeight" or "hidden"));
                    if (!skip)
                        writer.WriteAttributeString(reader.Prefix, reader.LocalName, reader.NamespaceURI, reader.Value);
                }
                while (reader.MoveToNextAttribute());
                reader.MoveToElement();
            }

            if (sized)
                WriteRowAttributes(writer, row);
            reader.Read();
        }
        else
        {
            XmlParts.CopyStartTag(reader, writer);
            if (!positioned)
                writer.WriteAttributeString("r", row.ToString(CultureInfo.InvariantCulture));
        }

        var sourceColumn = 0;
        if (!empty)
        {
            while (!(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
                {
                    var source = (XElement)XNode.ReadFrom(reader);
                    if (source.Attribute("r") is { } r && CellAddress.TryParse(r.Value, out var address))
                    {
                        if (address.Row != sourceRow)
                            throw new InvalidDataException($"Cell {address} of sheet '{_sheet.Name}' lies in row {sourceRow}; the file cannot be written back.");
                        sourceColumn = address.Column;
                    }
                    else
                    {
                        sourceColumn++;
                    }

                    var sourceAddress = new CellAddress(sourceRow, sourceColumn);
                    int column;
                    if (!_shifted)
                    {
                        column = sourceColumn;
                    }
                    else if (_map.MapColumn(_sheet, sourceColumn) is { } mappedColumn)
                    {
                        column = mappedColumn;
                    }
                    else
                    {
                        _changed = true;
                        continue;
                    }

                    CheckOrder(row, column);
                    NewCellsBefore(row, column, writer);
                    Cell(source, row, column, sourceAddress)?.WriteTo(writer);
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

    // Rows the file does not have: rows of new cells, and rows whose height or hidden flag changed.
    private void NewRowsBefore(int row, XmlWriter writer)
    {
        while (true)
        {
            var cellRow = _hasModel ? _model.Current.Row : int.MaxValue;
            var changedRow = _nextChangedRow < _changedRows.Count ? _changedRows[_nextChangedRow] : int.MaxValue;
            var current = Math.Min(cellRow, changedRow);
            if (current >= row || current == int.MaxValue)
                return;

            var cells = new List<XElement>();
            while (_hasModel && _model.Current.Row == current)
            {
                if (NewCell(_model.Current) is { } cell)
                    cells.Add(cell);
                Advance();
            }

            var attributes = changedRow == current;
            if (attributes)
                _nextChangedRow++;
            if (cells.Count == 0 && !attributes)
                continue;
            _changed = true;
            writer.WriteStartElement(_prefix, "row", _ns.NamespaceName);
            writer.WriteAttributeString("r", current.ToString(CultureInfo.InvariantCulture));
            if (attributes)
                WriteRowAttributes(writer, current);
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
    private static IEnumerable<ModelCell> ModelCells(Worksheet sheet, IEnumerable<CellAddress> more)
    {
        var extras = new List<CellAddress>();
        foreach (var address in more)
        {
            if (sheet.Store.Get(address.Row, address.Column) is null && sheet.StyleAt(address.Row, address.Column) is null
                && sheet.FormatAt(address.Row, address.Column) is null)
                extras.Add(address);
        }

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
    private XElement? NewCell(ModelCell cell) => Styled(NewContent(cell), null, cell.Row, cell.Column, null);

    private XElement? NewContent(ModelCell cell)
    {
        var address = new CellAddress(cell.Row, cell.Column);
        if (_headers.TryGetValue(address, out var header))
            return Rewrite(new XElement(_ns + "c", new XAttribute("r", address.ToString())), CellValue.Text(header), isFormula: false, null);
        if (cell.Data is not { } data)
            return null;
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
    private XElement? Cell(XElement source, int row, int column, CellAddress sourceAddress)
    {
        // The cell's position as the sheet has it now.
        var address = new CellAddress(row, column).ToString();
        if (source.Attribute("r") is { } r)
            r.Value = address;
        else
            source.ReplaceAttributes(new XAttribute("r", address), source.Attributes().ToList());
        return Styled(_headers.TryGetValue(new CellAddress(row, column), out var header)
                ? HeaderCell(source, row, column, header)
                : Content(source, row, column, sourceAddress),
            StyleIndex(source), row, column, sourceAddress);
    }

    // A new table's header cell holds its column's name as text.
    private XElement HeaderCell(XElement source, int row, int column, string name)
    {
        if (_hasModel && _model.Current.Row == row && _model.Current.Column == column)
            Advance();
        source.Elements(_ns + "f").Remove();
        source.Attribute("cm")?.Remove();
        return Rewrite(source, CellValue.Text(name), isFormula: false, null);
    }

    private static int? StyleIndex(XElement cell) =>
        int.TryParse((string?)cell.Attribute("s"), NumberStyles.None, CultureInfo.InvariantCulture, out var index) ? index : null;

    // The cell with the style index its look and number format need now; a cell left with neither
    // content nor style goes away.
    private XElement? Styled(XElement? element, int? sourceStyle, int row, int column, CellAddress? sourceAddress)
    {
        var target = TargetStyle(sourceStyle, row, column, sourceAddress);
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
    private int? TargetStyle(int? sourceStyle, int row, int column, CellAddress? sourceAddress)
    {
        var style = _sheet.StyleAt(row, column) ?? _sheet.Workbook.DefaultStyle;
        var code = _sheet.FormatAt(row, column)?.FormatCode ?? "General";
        var loadedStyle = sourceAddress is { } at && _source.Styles.TryGetValue(at, out var own) ? own : _plan.Source.Look.DefaultStyle;
        var loadedCode = sourceAddress is { } place && _source.Formats.TryGetValue(place, out var format) ? format : "General";
        if (style == loadedStyle && code == loadedCode)
            return sourceStyle;
        var index = _styles.XfFor(sourceStyle ?? 0, style, code);
        return index == 0 && sourceStyle is null ? null : index;
    }

    // The source cell with the model's content, or null when it has none.
    private XElement? Content(XElement source, int row, int column, CellAddress sourceAddress)
    {
        var address = new CellAddress(row, column);

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
        if (_shifted && saved.IsDataTable)
            throw new NotSupportedException($"{_sheet.Name}!{address} is part of an Excel data table, which SharpCell cannot move; save before inserting or deleting rows or columns there.");
        // A shared formula: si names its group, and the group's first cell (with ref) holds the text.
        var group = (string?)sourceFormula?.Attribute("t") == "shared" ? (string?)sourceFormula!.Attribute("si") : null;
        var holdsGroup = group is not null && sourceFormula!.Attribute("ref") is not null;
        if (data?.FormulaText is not null)
        {
            var loaded = !_shifted && WritePlan.IsLoadedFormula(data, _source, sourceAddress);
            if (_plan.KeptCells.Contains(new CellKey(_sheet, row, column)))
            {
                // SharpCell could not calculate it: the formula and the result stay as the file has them,
                // moved with their cell when rows or columns moved.
                Area? keptArea = null;
                if (sourceFormula is { } kept && (string?)kept.Attribute("t") == "array"
                    && (string?)kept.Attribute("ref") is { } reference && TryParseArea(reference, out var spill))
                {
                    keptArea = _shifted ? _map.Map(_sheet, spill) : spill;
                    if (keptArea is { } area)
                        _keptSpills.Add(area);
                }

                if (loaded)
                {
                    if (group is not null && !holdsGroup && _brokenGroups.Contains(group))
                        throw SharedGroupLost(address);
                    return source;
                }

                if (_shifted)
                    return WriteFormula(source, data, row, column, keepValue: true, keptArea);
            }

            // Changed formulas are written; so are the other cells of shared groups, which may have
            // lost the cell holding their text.
            if (!loaded || (_expandShared && group is not null && !holdsGroup))
            {
                if (holdsGroup && !loaded)
                    _brokenGroups.Add(group!);
                return WriteFormula(source, data, row, column);
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
        if (formulaRemoved && holdsGroup)
            _brokenGroups.Add(group!);

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

    private NotSupportedException SharedGroupLost(CellAddress address) =>
        new($"{_sheet.Name}!{address} shares the formula of a cell that changed, and SharpCell cannot read that formula to write it out.");

    // The cell with the model's formula in file syntax and its result; with keepValue, the result
    // the file has (for a formula SharpCell could not calculate), over keptArea if it spilled.
    private XElement WriteFormula(XElement cell, CellData data, int row, int column, bool keepValue = false, Area? keptArea = null)
    {
        cell.Elements(_ns + "f").Remove();
        cell.Attribute("cm")?.Remove();
        var origin = new CellAddress(row, column);
        var text = FormulaPrinter.PrintForFile(data.Formula!, origin, at => _sheet.Workbook.TableAt(_sheet, at.Row, at.Column)?.Name,
            name => FunctionValues.IsFunction(_sheet.Workbook, _sheet, name));
        var formula = new XElement(_ns + "f", text);
        if (data.FixedArray is { } fixedArea)
        {
            formula.Add(new XAttribute("t", "array"), new XAttribute("ref", Format(fixedArea)));
        }
        else if (!data.IsLegacy
                 && (keptArea is not null || data.SpillArea is not null || data.SpillWatch is not null || !FormulaStorage.IsPlain(data.Formula!, _sheet.Workbook, _sheet)))
        {
            // A dynamic array formula: its ref is the area it spills over, or just its cell.
            formula.Add(new XAttribute("t", "array"), new XAttribute("ref", Format(keptArea ?? data.SpillArea ?? Area.Cell(row, column))));
            cell.SetAttributeValue("cm", _metadata.DynamicArrayCm());
        }

        cell.AddFirst(formula);
        if (keepValue)
        {
            _changed = true;
            return cell;
        }

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
