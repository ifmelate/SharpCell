using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            + "<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>"
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
            + "<Relationship Id=\"rId2\" Type=\"" + Relationships + "/theme\" Target=\"theme/theme1.xml\"/>"
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
            foreach (var (name, content) in Parts.Append(("xl/theme/theme1.xml", Theme(workbook.ThemeColors))))
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

    // The workbook's theme colours as a theme part, so that theme colours in its cells show as the
    // workbook says; the Office theme's fonts and effects with them. The part lists dk1 before lt1.
    private static string Theme(IReadOnlyList<CellColor> colors)
    {
        const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        string[] names = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
        int[] index = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11];
        var scheme = string.Concat(names.Select((name, i) => $"<a:{name}><a:srgbClr val=\"{colors[index[i]].ToString()[1..]}\"/></a:{name}>"));
        const string fill = "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>";
        const string line = "<a:ln w=\"6350\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln>";
        const string effect = "<a:effectStyle><a:effectLst/></a:effectStyle>";
        return Head
            + $"<a:theme xmlns:a=\"{A}\" name=\"Office Theme\"><a:themeElements>"
            + $"<a:clrScheme name=\"Office\">{scheme}</a:clrScheme>"
            + "<a:fontScheme name=\"Office\">"
            + "<a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>"
            + "<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme>"
            + "<a:fmtScheme name=\"Office\">"
            + $"<a:fillStyleLst>{fill}{fill}{fill}</a:fillStyleLst><a:lnStyleLst>{line}{line}{line}</a:lnStyleLst>"
            + $"<a:effectStyleLst>{effect}{effect}{effect}</a:effectStyleLst><a:bgFillStyleLst>{fill}{fill}{fill}</a:bgFillStyleLst>"
            + "</a:fmtScheme></a:themeElements></a:theme>";
    }
}
