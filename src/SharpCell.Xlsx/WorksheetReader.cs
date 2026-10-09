using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell.Xlsx;

/// <summary>Streams the cells of one worksheet part into a sheet.</summary>
internal static class WorksheetReader
{
    // A parsed shared formula, or why it could not be parsed; every cell of the group uses it.
    private sealed record SharedFormula(FormulaNode? Node, string Text, string? Reason);

    private struct CellXml
    {
        public string? Type;
        public int? Cm;
        public int? Vm;
        public string? FormulaType;
        public string? FormulaText;
        public string? FormulaRef;
        public string? SharedIndex;
        public string? Value;
        public string? InlineText;
    }

    public static void Read(Package package, string part, Worksheet sheet, IReadOnlyList<string> sharedStrings, CellMetadata? metadata)
    {
        var loader = new SheetLoader(sheet);
        var shared = new Dictionary<string, SharedFormula>(StringComparer.Ordinal);
        var dateSystem = sheet.Workbook.DateSystem;
        using var reader = package.OpenXml(part);
        var row = 0;
        var column = 0;
        reader.Read();
        while (!reader.EOF)
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                reader.Read();
                continue;
            }

            if (reader.LocalName == "row")
            {
                // Rows and cells may leave out their position; it then follows the previous one.
                row = reader.GetAttribute("r") is { } r ? ParseRow(r, part) : row + 1;
                column = 0;
                reader.Read();
                continue;
            }

            if (reader.LocalName != "c")
            {
                reader.Read();
                continue;
            }

            if (reader.GetAttribute("r") is { } address)
            {
                if (!CellAddress.TryParse(address, out var parsed))
                    throw new InvalidDataException($"Cell reference '{address}' in '{part}' is not valid.");
                row = parsed.Row;
                column = parsed.Column;
            }
            else
            {
                column++;
                if (row is < 1 or > CellAddress.MaxRow || column > CellAddress.MaxColumn)
                    throw new InvalidDataException($"A cell in '{part}' lies outside the sheet.");
            }

            var cell = ReadCell(reader);
            var origin = new CellAddress(row, column);
            if (sheet.Store.Get(row, column) is not null)
                throw new InvalidDataException($"Cell {sheet.Name}!{origin} appears twice in '{part}'.");
            var cached = CachedValue(cell, sharedStrings, dateSystem, sheet, origin);
            if (cached.IsError && cell.Vm is { } vm && metadata?.RichError(vm) is { } richError)
                cached = CellValue.Error(richError);
            if (cell.FormulaType != "dataTable" && (cell.FormulaText is { Length: > 0 } || cell.FormulaType == "shared"))
                LoadFormula(loader, cell, origin, cached, metadata, shared);
            else
                loader.SetValue(row, column, cached);
        }

        loader.Complete();
    }

    private static void LoadFormula(SheetLoader loader, CellXml cell, CellAddress origin, CellValue cached, CellMetadata? metadata,
        Dictionary<string, SharedFormula> shared)
    {
        SharedFormula formula;
        if (cell.FormulaType == "shared" && cell.SharedIndex is { } index)
        {
            if (cell.FormulaText is { Length: > 0 } text)
                shared[index] = formula = Parse(text, origin);
            else if (!shared.TryGetValue(index, out formula!))
                formula = new SharedFormula(null, "", $"Shared formula {index} has no master cell.");
        }
        else if (cell.FormulaText is { Length: > 0 } text)
        {
            formula = Parse(text, origin);
        }
        else
        {
            loader.SetValue(origin.Row, origin.Column, cached);
            return;
        }

        if (formula.Node is null)
        {
            loader.SetUnsupportedFormula(origin.Row, origin.Column, formula.Text, formula.Reason!, cached);
            return;
        }

        var dynamic = cell.Cm is { } cm && (metadata?.IsDynamicArray(cm) ?? true);
        if (cell.FormulaType == "array")
        {
            var area = ParseArea(cell.FormulaRef, origin);

            // An array formula fills its whole area on every calculation, so its area counts against
            // the workbook's spill budget; one that does not fit is not calculated.
            if (!loader.SetFormula(origin.Row, origin.Column, formula.Node,
                    dynamic ? LoadedFormulaKind.Dynamic : LoadedFormulaKind.Array, area, cached))
            {
                loader.SetUnsupportedFormula(origin.Row, origin.Column, formula.Text,
                    $"The array formula covers {area.CellCount} cells, more than this workbook's spill budget leaves.", cached);
            }
        }
        else
        {
            loader.SetFormula(origin.Row, origin.Column, formula.Node,
                dynamic ? LoadedFormulaKind.Dynamic : LoadedFormulaKind.Legacy, null, cached);
        }
    }

    private static SharedFormula Parse(string text, CellAddress origin)
    {
        if (FormulaText.ReferencesOtherWorkbook(text))
            return new SharedFormula(null, text, FormulaText.ExternalReason);
        try
        {
            return new SharedFormula(FormulaParser.Parse("=" + text, origin), text, null);
        }
        catch (FormulaParseException ex)
        {
            return new SharedFormula(null, text, FormulaText.ParseFailure(ex));
        }
    }

    // Reads one <c> element and leaves the reader after it.
    private static CellXml ReadCell(XmlReader reader)
    {
        var cell = new CellXml
        {
            Type = reader.GetAttribute("t"),
            Cm = int.TryParse(reader.GetAttribute("cm"), NumberStyles.None, CultureInfo.InvariantCulture, out var cm) ? cm : null,
            Vm = int.TryParse(reader.GetAttribute("vm"), NumberStyles.None, CultureInfo.InvariantCulture, out var vm) ? vm : null,
        };
        if (reader.IsEmptyElement)
        {
            reader.Read();
            return cell;
        }

        var depth = reader.Depth;
        reader.Read();
        while (!(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) && !reader.EOF)
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                reader.Read();
                continue;
            }

            switch (reader.LocalName)
            {
                case "f":
                    cell.FormulaType = reader.GetAttribute("t");
                    cell.FormulaRef = reader.GetAttribute("ref");
                    cell.SharedIndex = reader.GetAttribute("si");
                    cell.FormulaText = reader.ReadElementContentAsString();
                    break;
                case "v":
                    cell.Value = reader.ReadElementContentAsString();
                    break;
                case "is":
                    cell.InlineText = XmlText.ReadStringItem(reader);
                    break;
                default:
                    reader.Skip();
                    break;
            }
        }

        reader.Read();
        return cell;
    }

    private static CellValue CachedValue(CellXml cell, IReadOnlyList<string> sharedStrings, DateSystem dateSystem, Worksheet sheet, CellAddress origin)
    {
        if (cell.Type == "inlineStr")
            return cell.InlineText is { } inline ? CellValue.Text(inline) : CellValue.Empty;
        if (cell.Value is not { } text)
            return CellValue.Empty;

        switch (cell.Type)
        {
            case "s":
                if (int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var index) && index < sharedStrings.Count)
                    return CellValue.Text(sharedStrings[index]);
                throw Invalid(sheet, origin, $"shared string index '{text}' does not exist");
            case "str":
                return CellValue.Text(XmlText.Decode(text));
            case "b":
                return text.Trim() switch
                {
                    "1" or "true" => CellValue.True,
                    "0" or "false" => CellValue.False,
                    _ => throw Invalid(sheet, origin, $"'{text}' is not a boolean"),
                };
            case "e":
                // Errors newer than SharpCell knows (#BLOCKED!, #CONNECT!, ...) read as #VALUE!.
                return CellValue.Error(ErrorKinds.TryParse(text.Trim(), out var error) ? error : ErrorKind.Value);
            case "d":
                if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date))
                    return CellValue.Number(DateSerial.FromDateTime(date, dateSystem));
                throw Invalid(sheet, origin, $"'{text}' is not a date");
            default:
                if (text.Length == 0)
                    return CellValue.Empty;
                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                    return CellValue.Number(number);
                throw Invalid(sheet, origin, $"'{text}' is not a number");
        }
    }

    // The area of an array formula starts at its own cell; anything else is taken as just that cell.
    private static Area ParseArea(string? reference, CellAddress origin)
    {
        var single = Area.Cell(origin.Row, origin.Column);
        if (reference is null)
            return single;

        var colon = reference.IndexOf(':');
        var firstText = colon < 0 ? reference.AsSpan() : reference.AsSpan(0, colon);
        var lastText = colon < 0 ? reference.AsSpan() : reference.AsSpan(colon + 1);
        if (!CellAddress.TryParse(firstText, out var first) || !CellAddress.TryParse(lastText, out var last)
            || first.Row != origin.Row || first.Column != origin.Column || last.Row < first.Row || last.Column < first.Column)
            return single;

        return new Area(first.Row, first.Column, last.Row, last.Column);
    }

    private static int ParseRow(string text, string part)
    {
        if (int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var row) && row is >= 1 and <= CellAddress.MaxRow)
            return row;
        throw new InvalidDataException($"Row number '{text}' in '{part}' is not valid.");
    }

    private static InvalidDataException Invalid(Worksheet sheet, CellAddress origin, string problem) =>
        new($"Cell {sheet.Name}!{origin}: {problem}.");
}
