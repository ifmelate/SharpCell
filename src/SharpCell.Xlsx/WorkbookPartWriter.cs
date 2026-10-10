using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace SharpCell.Xlsx;

/// <summary>Changes to the workbook part (xl/workbook.xml).</summary>
internal static class WorkbookPartWriter
{
    // CT_Workbook: calcPr comes after whichever of these the part has.
    private static readonly string[] BeforeCalcPr =
        ["fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "externalReferences", "definedNames"];

    /// <summary>The part with calcPr/@fullCalcOnLoad set: Excel recalculates every formula when it opens the file.</summary>
    public static byte[] WithFullCalcOnLoad(Package package, string part)
    {
        XDocument document;
        using (var reader = package.OpenXml(part))
            document = XDocument.Load(reader);
        var root = document.Root ?? throw new InvalidDataException($"'{part}' has no root element.");
        var ns = root.Name.Namespace;
        var calcPr = root.Element(ns + "calcPr");
        if (calcPr is null)
        {
            calcPr = new XElement(ns + "calcPr");
            var previous = root.Elements().LastOrDefault(e => e.Name.Namespace == ns && Array.IndexOf(BeforeCalcPr, e.Name.LocalName) >= 0);
            if (previous is null)
                root.AddFirst(calcPr);
            else
                previous.AddAfterSelf(calcPr);
        }

        calcPr.SetAttributeValue("fullCalcOnLoad", "1");
        return XmlParts.Save(document);
    }
}
