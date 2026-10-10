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
    /// <param name="newSheets">Sheets the file does not have yet, with the relationship id of their part.</param>
    public static void Write(PackageEdit edit, WritePlan plan, IReadOnlyList<(Worksheet Sheet, string RelationshipId)> newSheets)
    {
        var workbook = plan.Workbook;
        var look = plan.Source.Look;
        var part = plan.Source.WorkbookPart;
        var namesChanged = look.NamesChanged(workbook) || plan.Edits.Any;
        var iterationChanged = !workbook.Iteration.Equals(look.Iteration);
        if (newSheets.Count == 0 && !namesChanged && !iterationChanged && workbook.DateSystem == look.DateSystem && !plan.FullCalcOnLoad)
            return;

        var document = edit.ReadXml(part) ?? throw new System.IO.InvalidDataException($"The package has no part '{part}'.");
        var root = document.Root!;
        var ns = root.Name.Namespace;
        if (workbook.DateSystem != look.DateSystem)
            Section(root, "workbookPr").SetAttributeValue("date1904", workbook.DateSystem == DateSystem.Date1904 ? "1" : null);

        var sheets = Section(root, "sheets");
        if (newSheets.Count > 0)
        {
            var r = sheets.Elements(ns + "sheet").SelectMany(s => s.Attributes()).FirstOrDefault(a => a.Name.LocalName == "id" && a.Name.Namespace != XNamespace.None)?.Name.Namespace
                ?? TransitionalRelationships;
            var next = sheets.Elements(ns + "sheet").Select(s => (int?)s.Attribute("sheetId") ?? 0).DefaultIfEmpty(0).Max() + 1;
            foreach (var (sheet, id) in newSheets)
                sheets.Add(new XElement(ns + "sheet", new XAttribute("name", sheet.Name), new XAttribute("sheetId", next++), new XAttribute(r + "id", id)));
        }

        if (namesChanged)
            WriteNames(root, sheets, workbook, look, plan.Edits);

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

    // Names the file has keep their element, with the model's formula where it changed; names
    // SharpCell does not read (print areas, macro names) stay as they are; new names go at the end.
    private static void WriteNames(XElement root, XElement sheets, Workbook workbook, WorkbookLook look, SheetEditMap edits)
    {
        var ns = root.Name.Namespace;
        var sheetNames = sheets.Elements(ns + "sheet").Select(s => (string?)s.Attribute("name") ?? "").ToList();
        var section = root.Element(ns + "definedNames");
        var written = new HashSet<(string, string)>();
        var model = workbook.Names.All.ToDictionary(WorkbookLook.Key);
        foreach (var element in section?.Elements(ns + "definedName") ?? [])
        {
            var name = ((string?)element.Attribute("name"))?.ToUpperInvariant();
            if (name is null)
                continue;
            var scope = (int?)element.Attribute("localSheetId") is { } index && index >= 0 && index < sheetNames.Count ? sheetNames[index].ToUpperInvariant() : "";
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
                element.Value = FileText(definition);
        }

        foreach (var (key, definition) in model)
        {
            if (written.Contains(key))
                continue;
            section ??= Section(root, "definedNames");
            var element = new XElement(ns + "definedName", new XAttribute("name", definition.Spelling));
            if (definition.Scope is { } sheet)
                element.Add(new XAttribute("localSheetId", sheetNames.FindIndex(n => n.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase))));
            element.Add(FileText(definition));
            section.Add(element);
        }
    }

    private static string FileText(NameDefinition name) =>
        name.Formula is UnsupportedNode unsupported ? unsupported.Text : FormulaPrinter.PrintForFile(name.Formula, new CellAddress(1, 1), _ => null);

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
