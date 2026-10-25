using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using SharpCell.Parsing;

namespace SharpCell.Xlsx;

/// <summary>
/// Changes to the workbook part (xl/workbook.xml): new sheets, defined names, the date system,
/// iteration settings and the request to recalculate on opening. Everything else is kept.
/// </summary>
internal static class WorkbookWriter
{
    private static readonly XNamespace TransitionalRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // CT_Workbook's children in schema order.
    private static readonly string[] Order =
    [
        "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups",
        "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
        "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst",
    ];

    /// <param name="edit">The package edit the part goes into.</param>
    /// <param name="plan">The save.</param>
    /// <param name="parts">The worksheet part of every sheet of the workbook.</param>
    /// <param name="newSheets">Sheets the file does not have yet, with the relationship id of their part.</param>
    public static void Write(PackageEdit edit, WritePlan plan, IReadOnlyDictionary<Worksheet, string> parts,
        IReadOnlyList<(Worksheet Sheet, string RelationshipId)> newSheets)
    {
        var workbook = plan.Workbook;
        var look = plan.Source.Look;
        var part = plan.Source.WorkbookPart;
        var printChanged = workbook.Sheets.Where(sheet => PrintChanged(plan, sheet)).ToHashSet();
        var filterChanged = workbook.Sheets.Where(sheet => plan.Source.SheetFor(sheet).Look.Mapped(plan.Edits, sheet).FilterArea != sheet.AutoFilter?.Area).ToHashSet();
        var namesChanged = look.NamesChanged(workbook) || plan.Edits.Any || printChanged.Count > 0 || filterChanged.Count > 0;
        var iterationChanged = !workbook.Iteration.Equals(look.Iteration);
        var protectionChanged = look.ProtectionChanged(workbook);
        var sheetsChanged = newSheets.Count > 0 || SheetsChanged(plan, parts);
        if (!sheetsChanged && !namesChanged && !iterationChanged && !protectionChanged && workbook.DateSystem == look.DateSystem && !plan.FullCalcOnLoad)
            return;

        var document = edit.ReadXml(part) ?? throw new System.IO.InvalidDataException($"The package has no part '{part}'.");
        var root = document.Root!;
        var ns = root.Name.Namespace;
        if (workbook.DateSystem != look.DateSystem)
            Section(root, "workbookPr").SetAttributeValue("date1904", workbook.DateSystem == DateSystem.Date1904 ? "1" : null);

        if (protectionChanged)
            WriteProtection(root, workbook);

        var sheets = Section(root, "sheets");
        Dictionary<int, int>? indexes = null;
        if (sheetsChanged)
        {
            indexes = WriteSheets(edit, plan, sheets, parts, newSheets);
            WriteViews(root, sheets, indexes);
            ForgetSheetTitles(edit);
        }

        if (namesChanged || indexes is not null)
            WriteNames(root, sheets, workbook, look, plan.Edits, printChanged, filterChanged, indexes);

        if (iterationChanged || plan.FullCalcOnLoad)
        {
            var calcPr = Section(root, "calcPr");
            if (iterationChanged)
            {
                var iteration = workbook.Iteration;
                calcPr.SetAttributeValue("iterate", iteration.Enabled ? "1" : null);
                calcPr.SetAttributeValue("iterateCount", iteration.MaxIterations == 100 ? null : iteration.MaxIterations.ToString(CultureInfo.InvariantCulture));
                calcPr.SetAttributeValue("iterateDelta", iteration.MaxChange == 0.001 ? null : iteration.MaxChange.ToString("R", CultureInfo.InvariantCulture));
            }

            if (plan.FullCalcOnLoad)
                calcPr.SetAttributeValue("fullCalcOnLoad", "1");
        }

        edit.Parts[part] = XmlParts.Save(document);
    }

    // Whether a sheet of the file was removed, renamed, moved or shown or hidden.
    private static bool SheetsChanged(WritePlan plan, IReadOnlyDictionary<Worksheet, string> parts)
    {
        var live = new HashSet<string>(parts.Values, StringComparer.OrdinalIgnoreCase);
        var fileParts = plan.Source.FileSheets.Where(sheet => sheet.Part is not null).Select(sheet => sheet.Part!).ToList();
        if (fileParts.Any(part => !live.Contains(part)))
            return true;
        foreach (var (sheet, source) in plan.Source.LiveSheets())
        {
            if (sheet.Name != source.Name || sheet.Visibility != source.Look.Visibility)
                return true;
        }

        var order = plan.Workbook.Sheets.Select(sheet => parts[sheet]).Where(part => fileParts.Contains(part, StringComparer.OrdinalIgnoreCase));
        return !order.SequenceEqual(fileParts, StringComparer.OrdinalIgnoreCase);
    }

    // The sheet entries in the workbook's order, with their names and visibility; removed sheets go
    // with their parts. Chart sheets keep their places, worksheets fill theirs in the workbook's
    // order, and new sheets go at the end. Returns where each entry of the file went, by index.
    private static Dictionary<int, int> WriteSheets(PackageEdit edit, WritePlan plan, XElement sheets, IReadOnlyDictionary<Worksheet, string> parts,
        IReadOnlyList<(Worksheet Sheet, string RelationshipId)> newSheets)
    {
        var ns = sheets.Name.Namespace;
        var r = sheets.Elements(ns + "sheet").SelectMany(s => s.Attributes()).FirstOrDefault(a => a.Name.LocalName == "id" && a.Name.Namespace != XNamespace.None)?.Name.Namespace
            ?? TransitionalRelationships;
        var workbookPart = plan.Source.WorkbookPart;
        var relationships = edit.Package.ReadRelationships(workbookPart);
        var byPart = parts.ToDictionary(pair => pair.Value, pair => pair.Key, StringComparer.OrdinalIgnoreCase);
        var original = sheets.Elements(ns + "sheet").ToList();
        var elements = new Dictionary<Worksheet, XElement>();
        var slots = new List<(XElement Element, bool Worksheet, bool Removed)>();
        foreach (var element in original)
        {
            var id = (string?)element.Attribute(r + "id");
            if (id is null || !relationships.TryGetValue(id, out var relationship) || !relationship.Is("worksheet"))
            {
                slots.Add((element, false, false));
                continue;
            }

            if (byPart.TryGetValue(relationship.Target, out var sheet))
            {
                elements[sheet] = element;
                slots.Add((element, true, false));
                continue;
            }

            slots.Add((element, true, true));
            if (edit.Package.Exists(relationship.Target))
                edit.RemoveTree(workbookPart, relationship.Target);
            else
                edit.RemoveRelationship(workbookPart, id);
        }

        var next = original.Select(s => (int?)s.Attribute("sheetId") ?? 0).DefaultIfEmpty(0).Max() + 1;
        foreach (var (sheet, id) in newSheets)
            elements[sheet] = new XElement(ns + "sheet", new XAttribute("name", sheet.Name), new XAttribute("sheetId", next++), new XAttribute(r + "id", id));

        foreach (var sheet in plan.Workbook.Sheets)
        {
            var element = elements[sheet];
            if ((string?)element.Attribute("name") != sheet.Name)
                element.SetAttributeValue("name", sheet.Name);
            if (StateOf(element) != sheet.Visibility)
            {
                element.SetAttributeValue("state", sheet.Visibility switch
                {
                    SheetVisibility.Hidden => "hidden",
                    SheetVisibility.VeryHidden => "veryHidden",
                    _ => null,
                });
            }
        }

        var queue = new Queue<XElement>(plan.Workbook.Sheets.Select(sheet => elements[sheet]));
        var ordered = new List<XElement>();
        foreach (var (element, worksheet, removed) in slots)
        {
            if (!worksheet)
                ordered.Add(element);
            else if (!removed)
                ordered.Add(queue.Dequeue());
        }

        ordered.AddRange(queue);
        sheets.ReplaceNodes(ordered);

        var indexes = new Dictionary<int, int>();
        for (var i = 0; i < original.Count; i++)
        {
            var moved = ordered.IndexOf(original[i]);
            if (moved >= 0)
                indexes[i] = moved;
        }

        return indexes;
    }

    // The document properties list the sheets' names for file browsers; a list that no longer fits
    // goes (it is optional), and Excel writes it again when it saves the file.
    private static void ForgetSheetTitles(PackageEdit edit)
    {
        var properties = edit.Package.ReadRelationships("").Values.FirstOrDefault(r => r.Type.EndsWith("/extended-properties", StringComparison.Ordinal))?.Target;
        if (properties is null || edit.ReadXml(properties)?.Root is not { } root)
            return;
        var lists = root.Elements().Where(e => e.Name.LocalName is "HeadingPairs" or "TitlesOfParts").ToList();
        if (lists.Count == 0)
            return;
        lists.Remove();
        edit.Parts[properties] = XmlParts.Save(root.Document!);
    }

    private static SheetVisibility StateOf(XElement sheet) => (string?)sheet.Attribute("state") switch
    {
        "hidden" => SheetVisibility.Hidden,
        "veryHidden" => SheetVisibility.VeryHidden,
        _ => SheetVisibility.Visible,
    };

    // The workbook's windows show the sheet they showed, wherever it went. When it was removed or
    // hidden, the next visible sheet is shown, as in Excel.
    private static void WriteViews(XElement root, XElement sheets, Dictionary<int, int> indexes)
    {
        var ns = root.Name.Namespace;
        var entries = sheets.Elements(ns + "sheet").ToList();
        if (entries.Count == 0)
            return;
        foreach (var view in root.Element(ns + "bookViews")?.Elements(ns + "workbookView") ?? [])
        {
            var active = Moved((int?)view.Attribute("activeTab") ?? 0);
            if (StateOf(entries[active]) != SheetVisibility.Visible)
            {
                var visible = entries.FindIndex(active, e => StateOf(e) == SheetVisibility.Visible);
                active = visible >= 0 ? visible : Math.Max(0, entries.FindLastIndex(e => StateOf(e) == SheetVisibility.Visible));
            }

            view.SetAttributeValue("activeTab", active == 0 ? null : active.ToString(CultureInfo.InvariantCulture));
            var first = Math.Min(Moved((int?)view.Attribute("firstSheet") ?? 0), active);
            view.SetAttributeValue("firstSheet", first == 0 ? null : first.ToString(CultureInfo.InvariantCulture));
        }

        // Where an entry of the file went; a removed one's place goes to the next entry left.
        int Moved(int index)
        {
            for (var i = index; i < indexes.Keys.DefaultIfEmpty(-1).Max() + 1; i++)
            {
                if (indexes.TryGetValue(i, out var moved))
                    return moved;
            }

            return entries.Count - 1;
        }
    }

    // The locks and the password are rewritten on the element the file has, which keeps its other
    // attributes (revision protection); an element left with none is removed.
    private static void WriteProtection(XElement root, Workbook workbook)
    {
        var element = root.Element(root.Name.Namespace + "workbookProtection");
        var isProtected = workbook.IsStructureProtected || workbook.IsWindowsProtected || workbook.ProtectionPassword is not null;
        if (element is null)
        {
            if (!isProtected)
                return;
            element = Section(root, "workbookProtection");
        }

        var names = PasswordAttributes.Workbook.Names;
        element.Attributes().Where(a => a.Name.Namespace == XNamespace.None && (names.Contains(a.Name.LocalName)
            || a.Name.LocalName is "lockStructure" or "lockWindows")).Remove();
        element.Add(PasswordAttributes.Workbook.Write(workbook.ProtectionPassword));
        if (workbook.IsStructureProtected)
            element.SetAttributeValue("lockStructure", "1");
        if (workbook.IsWindowsProtected)
            element.SetAttributeValue("lockWindows", "1");
        if (!element.HasAttributes)
            element.Remove();
    }

    // Names the file has keep their element, with the model's formula where it changed; names
    // SharpCell does not read (print areas, macro names) stay as they are; new names go at the end.
    private static bool PrintChanged(WritePlan plan, Worksheet sheet)
    {
        // The model moves print areas and titles with inserted and deleted rows, also those listing
        // several ranges, which the formulas of other names do not follow.
        var look = plan.Source.SheetFor(sheet).Look.Mapped(plan.Edits, sheet);
        return look.PrintAreaChanged(sheet) || look.PrintTitlesChanged(sheet)
            || (plan.Edits.HasEdits(sheet) && (sheet.PrintAreas.Count > 0 || sheet.TitleRows is not null || sheet.TitleColumns is not null));
    }

    private static void WriteNames(XElement root, XElement sheets, Workbook workbook, WorkbookLook look, SheetEditMap edits, HashSet<Worksheet> printChanged,
        HashSet<Worksheet> filterChanged, Dictionary<int, int>? indexes)
    {
        var ns = root.Name.Namespace;
        var sheetNames = sheets.Elements(ns + "sheet").Select(s => (string?)s.Attribute("name") ?? "").ToList();
        var section = root.Element(ns + "definedNames");

        // Names local to a sheet follow it to its new place; those of a removed sheet go with it.
        if (indexes is not null)
        {
            foreach (var element in section?.Elements(ns + "definedName").ToList() ?? [])
            {
                if ((int?)element.Attribute("localSheetId") is not { } index)
                    continue;
                if (indexes.TryGetValue(index, out var moved))
                    element.SetAttributeValue("localSheetId", moved);
                else
                    element.Remove();
            }
        }

        var written = new HashSet<(string, string)>();
        var model = workbook.Names.All.ToDictionary(WorkbookLook.Key);
        var printWritten = new HashSet<(Worksheet, string)>();
        var databaseWritten = new HashSet<Worksheet>();
        foreach (var element in section?.Elements(ns + "definedName").ToList() ?? [])
        {
            var name = ((string?)element.Attribute("name"))?.ToUpperInvariant();
            if (name is null)
                continue;
            var scope = (int?)element.Attribute("localSheetId") is { } index && index >= 0 && index < sheetNames.Count ? sheetNames[index].ToUpperInvariant() : "";

            // The range of a filter the model changed: the model's, or none.
            if (name.Equals(FilterElements.DatabaseName, StringComparison.OrdinalIgnoreCase) && scope.Length > 0
                && workbook.TryGetSheet(scope, out var filtered) && filterChanged.Contains(filtered!))
            {
                if (FilterElements.DatabaseFormula(filtered!) is { } database && databaseWritten.Add(filtered!))
                    element.Value = database;
                else
                    element.Remove();
                continue;
            }

            // A print area or titles the model changed: the model's, or none.
            if (PrintNames.IsPrintName(name) && scope.Length > 0 && workbook.TryGetSheet(scope, out var printed) && printChanged.Contains(printed!))
            {
                var formula = PrintNames.IsArea(name) ? PrintNames.AreaFormula(printed!) : PrintNames.TitlesFormula(printed!);
                if (formula is null || !printWritten.Add((printed!, PrintNames.IsArea(name) ? PrintNames.Area : PrintNames.Titles)))
                    element.Remove();
                else
                    element.Value = formula;
                continue;
            }

            var key = model.Keys.FirstOrDefault(k => k.Name == name && k.Scope.ToUpperInvariant() == scope);
            if (key.Name is null || !written.Add(key))
            {
                // A name SharpCell does not read (a print area, a filter's range) follows moved rows too.
                var sheet = scope.Length > 0 && workbook.TryGetSheet(scope, out var scoped) ? scoped : null;
                element.Value = edits.MapFormula(element.Value, sheet, new CellAddress(1, 1));
                continue;
            }
            var definition = model[key];
            if (!look.Names.TryGetValue(key, out var text) || text != definition.Text)
                element.Value = FileText(definition, workbook);
        }

        foreach (var sheet in printChanged)
        {
            foreach (var (name, formula) in new[] { (PrintNames.Area, PrintNames.AreaFormula(sheet)), (PrintNames.Titles, PrintNames.TitlesFormula(sheet)) })
            {
                if (formula is null || printWritten.Contains((sheet, name)))
                    continue;
                section ??= Section(root, "definedNames");
                section.Add(new XElement(ns + "definedName", new XAttribute("name", name),
                    new XAttribute("localSheetId", sheetNames.FindIndex(n => n.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase))), formula));
            }
        }

        foreach (var sheet in filterChanged)
        {
            if (databaseWritten.Contains(sheet) || FilterElements.DatabaseFormula(sheet) is not { } database)
                continue;
            section ??= Section(root, "definedNames");
            section.Add(new XElement(ns + "definedName", new XAttribute("name", FilterElements.DatabaseName),
                new XAttribute("localSheetId", sheetNames.FindIndex(n => n.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase))),
                new XAttribute("hidden", "1"), database));
        }

        foreach (var (key, definition) in model)
        {
            if (written.Contains(key))
                continue;
            section ??= Section(root, "definedNames");
            var element = new XElement(ns + "definedName", new XAttribute("name", definition.Spelling));
            if (definition.Scope is { } sheet)
                element.Add(new XAttribute("localSheetId", sheetNames.FindIndex(n => n.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase))));
            element.Add(FileText(definition, workbook));
            section.Add(element);
        }
    }

    private static string FileText(NameDefinition name, Workbook workbook) =>
        name.Formula is UnsupportedNode unsupported
            ? unsupported.Text
            : FormulaPrinter.PrintForFile(name.Formula, new CellAddress(1, 1), _ => null, n => FunctionValues.IsFunction(workbook, name.Scope, n));

    // A child of the workbook, created where the schema puts it when the part has none.
    private static XElement Section(XElement root, string name)
    {
        var ns = root.Name.Namespace;
        if (root.Element(ns + name) is { } section)
            return section;
        section = new XElement(ns + name);
        var rank = Array.IndexOf(Order, name);
        var previous = root.Elements().LastOrDefault(e => e.Name.Namespace == ns && Array.IndexOf(Order, e.Name.LocalName) is var r && r >= 0 && r < rank);
        if (previous is null)
            root.AddFirst(section);
        else
            previous.AddAfterSelf(section);
        return section;
    }
}
