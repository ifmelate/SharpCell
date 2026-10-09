using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml;

namespace SharpCell.Xlsx;

/// <summary>
/// Reads .xlsx workbooks: sheets, values, formulas with the results Excel cached, defined names and
/// the date system. Styles, charts, pivot tables and macros are not read.
/// <para>
/// Every formula is out of date after loading: <see cref="Cell.Value"/> shows the value cached in
/// the file until <see cref="Workbook.Recalculate"/> calculates it. A formula SharpCell cannot
/// parse (for example a link to another workbook) does not fail the load; once calculated it is
/// <c>#NAME?</c> and listed in <see cref="Workbook.Diagnostics"/>.
/// </para>
/// </summary>
public static class XlsxReader
{
    /// <exception cref="InvalidDataException">The file is not a well-formed .xlsx workbook.</exception>
    /// <exception cref="NotSupportedException">The file is .xls, .xlsm, .xlsb or encrypted.</exception>
    public static Workbook Load(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        using var stream = File.OpenRead(path);
        return Load(stream);
    }

    /// <summary>Reads a workbook from a stream, which is left open.</summary>
    /// <inheritdoc cref="Load(string)"/>
    public static Workbook Load(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        return Load(stream, XlsxLimits.Default);
    }

    internal static Workbook Load(Stream stream, XlsxLimits limits)
    {
        try
        {
            using var package = Package.Open(stream, limits);
            return WorkbookReader.Read(package);
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

    public static Workbook Read(Package package)
    {
        var root = package.ReadRelationships("");
        var workbookPart = FirstOfType(root, "officeDocument")?.Target
            ?? throw new InvalidDataException("The package has no workbook part.");
        CheckContentType(package, workbookPart);

        var (date1904, sheets, names) = ReadWorkbookPart(package, workbookPart);
        var relationships = package.ReadRelationships(workbookPart);

        var workbook = new Workbook();
        if (date1904)
            workbook.DateSystem = DateSystem.Date1904;

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
            DefineName(workbook, name, byIndex);

        var sharedStringsPart = FirstOfType(relationships, "sharedStrings")?.Target;
        var sharedStrings = sharedStringsPart is not null && package.Exists(sharedStringsPart)
            ? SharedStrings.Read(package, sharedStringsPart)
            : [];
        var metadata = CellMetadata.Read(package, FirstOfType(relationships, "sheetMetadata")?.Target);

        var arrayBudget = new ArrayBudget(package.Limits.MaxArrayFormulaCells);
        foreach (var entry in byIndex)
        {
            if (entry is { } sheet)
                WorksheetReader.Read(package, sheet.Part, sheet.Sheet, sharedStrings, metadata, arrayBudget);
        }

        return workbook;
    }

    private static void CheckContentType(Package package, string workbookPart)
    {
        var contentType = package.Exists("[Content_Types].xml") ? package.ContentType(workbookPart) : null;
        if (contentType is null)
            return;
        if (contentType.Contains("macroEnabled", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("The file is a macro-enabled workbook (.xlsm/.xltm); only .xlsx files can be read.");
        if (contentType.Contains("binary", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("The file is a binary workbook (.xlsb); only .xlsx files can be read.");
    }

    private static (bool Date1904, List<SheetEntry> Sheets, List<NameEntry> Names) ReadWorkbookPart(Package package, string part)
    {
        var date1904 = false;
        var sheets = new List<SheetEntry>();
        var names = new List<NameEntry>();
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

        return (date1904, sheets, names);
    }

    // Leaves the reader after the element.
    private static NameEntry? ReadName(XmlReader reader)
    {
        var name = reader.GetAttribute("name");
        var local = reader.GetAttribute("localSheetId");
        var isMacro = CellMetadata.IsTrue(reader.GetAttribute("function")) || CellMetadata.IsTrue(reader.GetAttribute("vbProcedure"))
            || CellMetadata.IsTrue(reader.GetAttribute("xlm"));
        var formula = reader.ReadElementContentAsString();

        // Built-in names (print areas, filters) and placeholders for newer functions are not user names.
        if (name is null || isMacro || formula.Length == 0
            || name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)
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
            workbook.DefineName(name.Name, name.Formula, scope);
        }
        catch (FormulaParseException ex)
        {
            workbook.DefineUnsupportedName(name.Name, name.Formula, FormulaText.ParseFailure(ex), scope);
        }
    }

    private static string? RelationshipId(XmlReader reader)
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

    private static Relationship? FirstOfType(Dictionary<string, Relationship> relationships, string kind)
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

/// <summary>Cells array formulas may still cover; shared by all sheets of a workbook.</summary>
internal sealed class ArrayBudget(long cells)
{
    public long Limit { get; } = cells;

    public long Remaining { get; private set; } = cells;

    public bool TryTake(long count)
    {
        if (count > Remaining)
            return false;
        Remaining -= count;
        return true;
    }
}

internal static class FormulaText
{
    public const string ExternalReason = "Links to other workbooks are not supported.";

    public static string ParseFailure(FormulaParseException ex) => $"The formula cannot be parsed: {ex.Message}";

    /// <summary>
    /// Whether the formula refers to another workbook, which files write as an index in brackets:
    /// <c>[1]Sheet1!A1</c>, <c>'[1]My sheet'!A1</c>, <c>[1]!Name</c>. Table references such as
    /// <c>Table1[Col]</c> follow a name, so they do not match.
    /// </summary>
    public static bool ReferencesOtherWorkbook(string formula)
    {
        var inString = false;
        for (var i = 0; i < formula.Length; i++)
        {
            var ch = formula[i];
            if (ch == '"')
            {
                inString = !inString;
                continue;
            }

            if (inString || ch != '[')
                continue;
            var previous = i == 0 ? ' ' : formula[i - 1];
            if (char.IsLetterOrDigit(previous) || previous is '_' or '.' or '[' or ']' or '\\')
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
