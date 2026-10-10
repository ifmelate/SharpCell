using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace SharpCell.Tests.Xlsx;

/// <summary>Builds minimal .xlsx packages in memory, part by part, for reader tests.</summary>
public sealed class TestXlsx
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private readonly List<(string Name, string Xml, string Type)> _sheets = [];
    private readonly Dictionary<string, string> _extraParts = [];

    public string WorkbookExtra { get; set; } = "";

    /// <summary>Markup after definedNames, such as calcPr.</summary>
    public string WorkbookTail { get; set; } = "";

    public string? SharedStrings { get; set; }

    public string? Metadata { get; set; }

    /// <summary>The whole xl/styles.xml part, or null for none.</summary>
    public string? Styles { get; set; }

    /// <summary>The whole xl/theme/theme1.xml part, or null for none.</summary>
    public string? Theme { get; set; }

    public string WorkbookContentType { get; set; } = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

    /// <summary>Use absolute targets ("/xl/...") in the workbook relationships.</summary>
    public bool AbsoluteTargets { get; set; }

    public TestXlsx Sheet(string name, string sheetData, string attributes = "")
    {
        _sheets.Add((name, $"<worksheet xmlns=\"{Main}\"><sheetData>{sheetData}</sheetData></worksheet>", "worksheet" + attributes));
        return this;
    }

    /// <summary>A worksheet with the given content inside &lt;worksheet&gt;, for elements other than sheetData.</summary>
    public TestXlsx RawSheet(string name, string content)
    {
        _sheets.Add((name, $"<worksheet xmlns=\"{Main}\" xmlns:r=\"{Rel}\">{content}</worksheet>", "worksheet"));
        return this;
    }

    public TestXlsx ChartSheet(string name)
    {
        _sheets.Add((name, $"<chartsheet xmlns=\"{Main}\"/>", "chartsheet"));
        return this;
    }

    public TestXlsx Part(string name, string content)
    {
        _extraParts[name] = content;
        return this;
    }

    public MemoryStream Build()
    {
        var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            var types = new StringBuilder("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            types.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            types.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            types.Append($"<Override PartName=\"/xl/workbook.xml\" ContentType=\"{WorkbookContentType}\"/></Types>");
            Write(zip, "[Content_Types].xml", types.ToString());
            Write(zip, "_rels/.rels",
                $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{Rel}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");

            var workbook = new StringBuilder($"<workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\">{WorkbookExtra}<sheets>");
            var rels = new StringBuilder("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            var prefix = AbsoluteTargets ? "/xl/" : "";
            for (var i = 0; i < _sheets.Count; i++)
            {
                var (name, xml, type) = _sheets[i];
                var kind = type.StartsWith("chartsheet") ? "chartsheet" : "worksheet";
                var state = type.Length > "worksheet".Length && kind == "worksheet" ? type["worksheet".Length..] : "";
                var part = $"{kind}s/sheet{i + 1}.xml";
                workbook.Append($"<sheet name=\"{name}\" sheetId=\"{i + 1}\"{state} r:id=\"rIdS{i + 1}\"/>");
                rels.Append($"<Relationship Id=\"rIdS{i + 1}\" Type=\"{Rel}/{kind}\" Target=\"{prefix}{part}\"/>");
                Write(zip, "xl/" + part, xml);
            }

            workbook.Append("</sheets>");
            if (_definedNames.Length > 0)
                workbook.Append("<definedNames>").Append(_definedNames).Append("</definedNames>");
            workbook.Append(WorkbookTail);
            workbook.Append("</workbook>");
            if (SharedStrings is not null)
            {
                rels.Append($"<Relationship Id=\"rIdSS\" Type=\"{Rel}/sharedStrings\" Target=\"{prefix}sharedStrings.xml\"/>");
                Write(zip, "xl/sharedStrings.xml", $"<sst xmlns=\"{Main}\">{SharedStrings}</sst>");
            }

            if (Metadata is not null)
            {
                rels.Append($"<Relationship Id=\"rIdMD\" Type=\"{Rel}/sheetMetadata\" Target=\"{prefix}metadata.xml\"/>");
                Write(zip, "xl/metadata.xml", Metadata);
            }

            if (Styles is not null)
            {
                rels.Append($"<Relationship Id=\"rIdST\" Type=\"{Rel}/styles\" Target=\"{prefix}styles.xml\"/>");
                Write(zip, "xl/styles.xml", Styles);
            }

            if (Theme is not null)
            {
                rels.Append($"<Relationship Id=\"rIdTH\" Type=\"{Rel}/theme\" Target=\"{prefix}theme/theme1.xml\"/>");
                Write(zip, "xl/theme/theme1.xml", Theme);
            }

            rels.Append("</Relationships>");
            Write(zip, "xl/workbook.xml", workbook.ToString());
            Write(zip, "xl/_rels/workbook.xml.rels", rels.ToString());
            foreach (var (name, content) in _extraParts)
                Write(zip, name, content);
        }

        stream.Position = 0;
        return stream;
    }

    private readonly StringBuilder _definedNames = new();

    public TestXlsx Name(string name, string formula, string attributes = "")
    {
        _definedNames.Append($"<definedName name=\"{name}\"{attributes}>{formula}</definedName>");
        return this;
    }

    /// <summary>The metadata part Excel writes for dynamic array formulas: cm="1" marks one.</summary>
    public const string DynamicArrayMetadata =
        "<metadata xmlns=\"" + Main + "\" xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\">"
        + "<metadataTypes count=\"1\"><metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\"/></metadataTypes>"
        + "<futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst><ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\">"
        + "<xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk></futureMetadata>"
        + "<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata></metadata>";

    private static void Write(ZipArchive zip, string name, string content)
    {
        var entry = zip.CreateEntry(name);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}
