using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace SharpCell.Xlsx;

/// <summary>
/// The package a workbook built in code is saved into: a workbook without sheets and the default
/// styles. The writer adds the sheets and everything else, as it does for a file that was read.
/// </summary>
internal static class Template
{
    /// <summary>The time Excel gives every entry of the files it writes; fixed, so a save is repeatable.</summary>
    public static readonly DateTimeOffset EntryTime = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string Head = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";

    private static readonly (string Name, string Content)[] Parts =
    [
        ("[Content_Types].xml", Head
            + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
            + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
            + "</Types>"),
        ("_rels/.rels", Head
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"" + Relationships + "/officeDocument\" Target=\"xl/workbook.xml\"/>"
            + "</Relationships>"),
        ("xl/workbook.xml", Head
            + "<workbook xmlns=\"" + Main + "\" xmlns:r=\"" + Relationships + "\">"
            + "<bookViews><workbookView/></bookViews><sheets/><calcPr calcId=\"191029\"/></workbook>"),
        ("xl/_rels/workbook.xml.rels", Head
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"" + Relationships + "/styles\" Target=\"styles.xml\"/>"
            + "</Relationships>"),
        ("xl/styles.xml", Head
            + "<styleSheet xmlns=\"" + Main + "\">"
            + "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>"
            + "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
            + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
            + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
            + "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>"
            + "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
            + "</styleSheet>"),
    ];

    /// <summary>A source for a workbook that was not read from a file: the template, with none of the workbook's sheets.</summary>
    public static XlsxSource SourceFor(Workbook workbook)
    {
        var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, content) in Parts)
            {
                var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
                entry.LastWriteTime = EntryTime;
                using var stream = entry.Open();
                stream.Write(Encoding.UTF8.GetBytes(content));
            }
        }

        var bytes = output.ToArray();
        var empty = XlsxReader.Load(bytes, XlsxLimits.Default);
        return new XlsxSource(bytes, XlsxLimits.Default, "xl/workbook.xml", new System.Collections.Generic.Dictionary<Worksheet, SheetSource>(),
            WorkbookLook.Of(empty, new System.Collections.Generic.Dictionary<Worksheet, string>()), []);
    }
}
