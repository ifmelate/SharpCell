using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml;
using SharpCell.Parsing;

namespace SharpCell.Xlsx;

/// <summary>
/// Reads .xlsx workbooks: sheets, values, formulas with the results Excel cached, defined names,
/// tables, hidden rows, whether a sheet is filtered, the date system, number formats, cell styles,
/// column widths, row heights, merged cells, frozen panes and gridlines, from .xlsx and macro-enabled
/// .xlsm files. Charts, pivot tables and macros are not read (a save keeps them; macros never run).
/// <para>
/// Every formula is out of date after loading: <see cref="Cell.Value"/> shows the value cached in
/// the file until <see cref="Workbook.Recalculate"/> calculates it. A formula SharpCell cannot
/// parse (for example a link to another workbook) does not fail the load; once calculated it is
/// <c>#NAME?</c> and listed in <see cref="Workbook.Diagnostics"/>.
/// </para>
/// <para>
/// The workbook keeps the file's bytes in memory, so <see cref="XlsxWriter"/> can write changes
/// into the same file.
/// </para>
/// </summary>
public static class XlsxReader
{
    /// <summary>Reads a workbook from a file.</summary>
    /// <param name="path">Path of the .xlsx file.</param>
    /// <returns>The workbook, with the results Excel saved until it is recalculated.</returns>
    /// <exception cref="InvalidDataException">The file is not a well-formed .xlsx workbook.</exception>
    /// <exception cref="NotSupportedException">The file is .xls, .xlsb or encrypted.</exception>
    public static Workbook Load(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        return Load(File.ReadAllBytes(path), XlsxLimits.Default);
    }

    /// <summary>Reads a workbook from a stream, which is left open.</summary>
    /// <param name="stream">The .xlsx content; it need not be seekable.</param>
    /// <inheritdoc cref="Load(string)"/>
    public static Workbook Load(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        return Load(stream, XlsxLimits.Default);
    }

    internal static Workbook Load(Stream stream, XlsxLimits limits)
    {
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return Load(copy.ToArray(), limits);
    }

    // The file stays in memory with the workbook: XlsxWriter writes changes into it.
    internal static Workbook Load(byte[] bytes, XlsxLimits limits)
    {
        try
        {
            using var package = Package.Open(new MemoryStream(bytes, writable: false), limits);
            var workbook = WorkbookReader.Read(package, out var workbookPart, out var sheetParts);
            workbook.Source = XlsxSource.Capture(bytes, limits, workbookPart, workbook, sheetParts);
            return workbook;
        }
        catch (XmlException ex)
        {
            throw new InvalidDataException($"The workbook contains malformed XML: {ex.Message}", ex);
        }
    }
}

internal static class WorkbookReader
{
    private sealed record SheetEntry(string Name, string RelationshipId);

    private sealed record NameEntry(string Name, int? LocalSheet, string Formula);

    public static Workbook Read(Package package, out string workbookPart, out Dictionary<Worksheet, string> sheetParts)
    {
        var root = package.ReadRelationships("");
        workbookPart = FirstOfType(root, "officeDocument")?.Target
            ?? throw new InvalidDataException("The package has no workbook part.");
        CheckContentType(package, workbookPart);

        var (date1904, iteration, sheets, names, protection) = ReadWorkbookPart(package, workbookPart);
        var relationships = package.ReadRelationships(workbookPart);

        var workbook = new Workbook { MaxSpillCells = package.Limits.MaxSpillCells };
        if (date1904)
            workbook.DateSystem = DateSystem.Date1904;
        workbook.Iteration = iteration;
        workbook.SetProtection(protection.Structure, protection.Windows, protection.Password);

        // Chart sheets and other non-worksheets hold no cells, but they count in sheet indexes.
        var byIndex = new List<(Worksheet Sheet, string Part)?>();
        foreach (var entry in sheets)
        {
            if (!relationships.TryGetValue(entry.RelationshipId, out var relationship) || !relationship.Is("worksheet"))
            {
                byIndex.Add(null);
                continue;
            }

            Worksheet sheet;
            try
            {
                sheet = workbook.AddSheet(entry.Name);
            }
            catch (ArgumentException ex)
            {
                throw new InvalidDataException($"Sheet name '{entry.Name}' is not valid: {ex.Message}", ex);
            }

            byIndex.Add((sheet, relationship.Target));
        }

        foreach (var name in names)
        {
            if (name.Name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase))
                PrintName(name, byIndex);
            else
                DefineName(workbook, name, byIndex);
        }

        var saved = ReadSavedCells(package, relationships);
        if (saved.Styles.Count > 0)
            workbook.DefaultStyle = saved.Styles[0].Style;

        sheetParts = new Dictionary<Worksheet, string>();
        foreach (var entry in byIndex)
        {
            if (entry is { } sheet)
            {
                WorksheetReader.Read(package, sheet.Part, sheet.Sheet, saved);
                NoteParts.Read(package, sheet.Part, sheet.Sheet);
                sheetParts[sheet.Sheet] = sheet.Part;
            }
        }

        return workbook;
    }

    private static void CheckContentType(Package package, string workbookPart)
    {
        var contentType = package.Exists("[Content_Types].xml") ? package.ContentType(workbookPart) : null;
        if (contentType is null)
            return;
        // Macro-enabled workbooks (.xlsm, .xltm) are read like any other: macros are kept, never run.
        if (contentType.Contains("binary", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("The file is a binary workbook (.xlsb); only .xlsx and .xlsm files can be read.");
    }

    private static (bool Date1904, IterationSettings Iteration, List<SheetEntry> Sheets, List<NameEntry> Names, (bool Structure, bool Windows, ProtectionPassword? Password) Protection) ReadWorkbookPart(Package package, string part)
    {
        var date1904 = false;
        var iteration = new IterationSettings();
        var sheets = new List<SheetEntry>();
        var names = new List<NameEntry>();
        var protection = (Structure: false, Windows: false, Password: (ProtectionPassword?)null);
        using var reader = package.OpenXml(part);
        reader.Read();
        while (!reader.EOF)
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                reader.Read();
                continue;
            }

            switch (reader.LocalName)
            {
                case "workbookPr":
                    date1904 = CellMetadata.IsTrue(reader.GetAttribute("date1904"));
                    break;
                case "workbookProtection":
                    protection = (CellMetadata.IsTrue(reader.GetAttribute("lockStructure")), CellMetadata.IsTrue(reader.GetAttribute("lockWindows")),
                        PasswordAttributes.Workbook.Read(reader.GetAttribute));
                    break;
                case "calcPr":
                    iteration = ReadIteration(reader);
                    break;
                case "sheet":
                    if (reader.GetAttribute("name") is { } sheetName && RelationshipId(reader) is { } id)
                        sheets.Add(new SheetEntry(sheetName, id));
                    break;
                case "definedName":
                    if (ReadName(reader) is { } name)
                        names.Add(name);
                    continue;
            }

            reader.Read();
        }

        return (date1904, iteration, sheets, names, protection);
    }

    // calcPr's defaults are Excel's: iterate false, iterateCount 100, iterateDelta 0.001.
    private static IterationSettings ReadIteration(XmlReader reader)
    {
        var count = 100;
        var delta = 0.001;
        if (reader.GetAttribute("iterateCount") is { } countText
            && !int.TryParse(countText, NumberStyles.Integer, CultureInfo.InvariantCulture, out count))
            throw new InvalidDataException($"calcPr iterateCount '{countText}' is not a number.");
        if (reader.GetAttribute("iterateDelta") is { } deltaText
            && !double.TryParse(deltaText, NumberStyles.Float, CultureInfo.InvariantCulture, out delta))
            throw new InvalidDataException($"calcPr iterateDelta '{deltaText}' is not a number.");
        try
        {
            return new IterationSettings(CellMetadata.IsTrue(reader.GetAttribute("iterate")), count, delta);
        }
        catch (ArgumentOutOfRangeException ex)
        {
            throw new InvalidDataException($"calcPr has iteration settings Excel does not allow: {ex.Message}", ex);
        }
    }

    // Leaves the reader after the element.
    private static NameEntry? ReadName(XmlReader reader)
    {
        var name = reader.GetAttribute("name");
        var local = reader.GetAttribute("localSheetId");
        var isMacro = CellMetadata.IsTrue(reader.GetAttribute("function")) || CellMetadata.IsTrue(reader.GetAttribute("vbProcedure"))
            || CellMetadata.IsTrue(reader.GetAttribute("xlm"));
        var formula = reader.ReadElementContentAsString();

        // Built-in names are not user names: print areas and titles are read as the sheet's own,
        // others (a filter's range) stay in the file; nor are placeholders for newer functions.
        if (name is null || isMacro || formula.Length == 0
            || (name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase) && !PrintNames.IsPrintName(name))
            || name.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase))
            return null;

        int? sheet = null;
        if (local is not null)
        {
            if (!int.TryParse(local, NumberStyles.None, CultureInfo.InvariantCulture, out var index))
                return null;
            sheet = index;
        }

        return new NameEntry(name, sheet, formula);
    }

    // A print area or print titles: the sheet's own setting. One SharpCell cannot read stays in the file.
    private static void PrintName(NameEntry name, List<(Worksheet Sheet, string Part)?> byIndex)
    {
        if (name.LocalSheet is not { } index || index >= byIndex.Count || byIndex[index] is not { } entry
            || PrintNames.Parse(name.Formula) is not { } areas)
            return;
        if (PrintNames.IsArea(name.Name))
        {
            entry.Sheet.LoadPrinting(areas, null, null);
            return;
        }

        (int, int)? rows = null, columns = null;
        foreach (var area in areas)
        {
            if (area.FirstColumn == 1 && area.LastColumn == CellAddress.MaxColumn)
                rows = (area.FirstRow, area.LastRow);
            else if (area.FirstRow == 1 && area.LastRow == CellAddress.MaxRow)
                columns = (area.FirstColumn, area.LastColumn);
        }

        entry.Sheet.LoadPrinting(null, rows, columns);
    }

    private static void DefineName(Workbook workbook, NameEntry name, List<(Worksheet Sheet, string Part)?> byIndex)
    {
        Worksheet? scope = null;
        if (name.LocalSheet is { } index)
        {
            if (index >= byIndex.Count || byIndex[index] is not { } entry)
                return;
            scope = entry.Sheet;
        }

        if (!Workbook.IsValidName(name.Name))
            return;
        if (FormulaText.ReferencesOtherWorkbook(name.Formula))
        {
            workbook.DefineUnsupportedName(name.Name, name.Formula, FormulaText.ExternalReason, scope);
            return;
        }

        try
        {
            // As cell formulas are: without the prefixes files give newer functions.
            var text = name.Formula.StartsWith('=') ? name.Formula : "=" + name.Formula;
            workbook.DefineName(name.Name, "=" + FormulaPrinter.Print(FormulaParser.Parse(text, new CellAddress(1, 1)), new CellAddress(1, 1)), scope);
        }
        catch (FormulaParseException ex)
        {
            workbook.DefineUnsupportedName(name.Name, name.Formula, FormulaText.ParseFailure(ex), scope);
        }
    }

    internal static string? RelationshipId(XmlReader reader)
    {
        if (!reader.MoveToFirstAttribute())
            return null;
        string? id = null;
        do
        {
            if (reader.LocalName == "id" && reader.NamespaceURI.EndsWith("/relationships", StringComparison.Ordinal))
                id = reader.Value;
        }
        while (reader.MoveToNextAttribute());
        reader.MoveToElement();
        return id;
    }

    /// <summary>The parts cells point into: shared strings, and the metadata behind dynamic arrays and newer errors.</summary>
    internal static SavedCells ReadSavedCells(Package package, Dictionary<string, Relationship> relationships)
    {
        var sharedStringsPart = FirstOfType(relationships, "sharedStrings")?.Target;
        var sharedStrings = sharedStringsPart is not null && package.Exists(sharedStringsPart)
            ? SharedStrings.Read(package, sharedStringsPart)
            : [];
        var rich = RichValues.Read(package,
            FirstOfType(relationships, "rdRichValue")?.Target ?? "xl/richData/rdrichvalue.xml",
            FirstOfType(relationships, "rdRichValueStructure")?.Target ?? "xl/richData/rdrichvaluestructure.xml");
        var metadata = CellMetadata.Read(package, FirstOfType(relationships, "sheetMetadata")?.Target, rich);
        var stylesPart = FirstOfType(relationships, "styles")?.Target;
        var theme = ThemeReader.Read(package, FirstOfType(relationships, "theme")?.Target);
        var styles = stylesPart is not null && package.Exists(stylesPart) ? StylesReader.Read(package, stylesPart, theme) : new StyleSheet();
        return new SavedCells(sharedStrings, metadata, styles.Xfs, styles.Dxfs);
    }

    internal static Relationship? FirstOfType(Dictionary<string, Relationship> relationships, string kind)
    {
        Relationship? first = null;
        foreach (var relationship in relationships.Values)
        {
            if (relationship.Is(kind) && (first is null || string.CompareOrdinal(relationship.Id, first.Id) < 0))
                first = relationship;
        }

        return first;
    }
}

internal static class FormulaText
{
    public const string ExternalReason = "Links to other workbooks are not supported.";

    public static string ParseFailure(FormulaParseException ex) => $"The formula cannot be parsed: {ex.Message}";

    /// <summary>
    /// Whether the formula refers to another workbook, which files write as an index in brackets:
    /// <c>[1]Sheet1!A1</c>, <c>'[1]My sheet'!A1</c>, <c>[1]!Name</c>. Table references such as
    /// <c>Table1[Col]</c> follow a name, and brackets inside them (<c>Table1[[#This Row],[1]]</c>)
    /// are nested, so they do not match.
    /// </summary>
    public static bool ReferencesOtherWorkbook(string formula)
    {
        var inString = false;
        var depth = 0;
        for (var i = 0; i < formula.Length; i++)
        {
            var ch = formula[i];
            if (ch == '"')
            {
                inString = !inString;
                continue;
            }

            if (inString)
                continue;
            if (ch == ']')
            {
                depth = Math.Max(0, depth - 1);
                continue;
            }

            if (ch != '[')
                continue;
            depth++;
            var previous = i == 0 ? ' ' : formula[i - 1];
            if (depth > 1 || char.IsLetterOrDigit(previous) || previous is '_' or '.' or '\\')
                continue;

            var j = i + 1;
            while (j < formula.Length && char.IsAsciiDigit(formula[j]))
                j++;
            if (j > i + 1 && j < formula.Length && formula[j] == ']')
                return true;
        }

        return false;
    }
}
