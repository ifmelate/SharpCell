using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace SharpCell.Xlsx;

/// <summary>Writes the file of a <see cref="WritePlan"/>: the source package with the parts that changed.</summary>
internal static class PackageWriter
{
    private const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

    private const string NewWorksheet =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
        + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData/></worksheet>";

    /// <summary>The file with the workbook's values; the source bytes themselves when nothing changed.</summary>
    public static byte[] Write(WritePlan plan)
    {
        var source = plan.Source;
        var workbook = plan.Workbook;
        using var package = Package.Open(new MemoryStream(source.Bytes, writable: false), source.Limits);
        var edit = new PackageEdit(package);
        var relationships = package.ReadRelationships(source.WorkbookPart);
        var saved = WorkbookReader.ReadSavedCells(package, relationships);
        var metadata = CellMetadataWriter.Open(edit, source.WorkbookPart, relationships);
        var styles = StylesWriter.Open(edit, source.WorkbookPart, relationships);
        if (workbook.DefaultStyle != source.Look.DefaultStyle)
            styles.SetDefault(workbook.DefaultStyle);

        var folder = source.WorkbookPart[..(source.WorkbookPart.LastIndexOf('/') + 1)];
        var parts = new Dictionary<Worksheet, string>();
        var newSheets = new List<(Worksheet Sheet, string RelationshipId)>();
        foreach (var sheet in workbook.Sheets)
        {
            if (source.Sheets.TryGetValue(sheet, out var known))
            {
                parts[sheet] = known.Part!;
                continue;
            }

            var part = edit.FreeName(folder + "worksheets/sheet{0}.xml");
            edit.Parts[part] = [];   // claims the name
            edit.SetContentType(part, WorksheetContentType);
            newSheets.Add((sheet, edit.AddRelationship(source.WorkbookPart, PackageEdit.OfficeRelationships + "worksheet", part)));
            parts[sheet] = part;
        }

        var tables = new Dictionary<Worksheet, List<string>>();
        var headers = new Dictionary<Worksheet, Dictionary<CellAddress, string>>();
        var tableIds = TableWriter.UsedIds(edit);
        var nextId = 1;
        foreach (var table in workbook.Tables)
        {
            if (source.Look.Tables.Contains(table.Name.ToUpperInvariant()))
                continue;
            while (!tableIds.Add(nextId))
                nextId++;
            var part = edit.FreeName(folder + "tables/table{0}.xml");
            edit.Parts[part] = XmlParts.Save(TableWriter.Write(table, nextId));
            edit.SetContentType(part, TableWriter.ContentType);
            var sheetPart = parts[table.Worksheet];
            if (!tables.TryGetValue(table.Worksheet, out var ids))
                tables[table.Worksheet] = ids = [];
            ids.Add(edit.AddRelationship(sheetPart, PackageEdit.OfficeRelationships + "table", part));
            if (!headers.TryGetValue(table.Worksheet, out var names))
                headers[table.Worksheet] = names = [];
            foreach (var (address, name) in TableWriter.HeaderCells(table))
                names[address] = name;
        }

        foreach (var sheet in workbook.Sheets)
        {
            var sheetSource = source.SheetFor(sheet);
            using var reader = sheetSource.Part is { } part
                ? package.OpenXml(part)
                : System.Xml.XmlReader.Create(new StringReader(NewWorksheet));
            var bytes = WorksheetWriter.Write(reader, sheet, sheetSource, plan, saved, metadata, styles,
                tables.TryGetValue(sheet, out var ids) ? ids : [], headers.TryGetValue(sheet, out var names) ? names : []);
            if (bytes is not null || sheetSource.Part is null)
                edit.Parts[parts[sheet]] = bytes ?? Encoding.UTF8.GetBytes(NewWorksheet);
        }

        metadata.Save();
        styles.Save();

        // Excel trusts the calculation chain; one that lists cells without formulas makes it repair the file.
        if (plan.AnyFormulaChanged && WorkbookReader.FirstOfType(relationships, "calcChain")?.Target is { } chain)
            edit.RemovePart(source.WorkbookPart, chain);

        WorkbookWriter.Write(edit, plan, newSheets);
        edit.Complete();
        return edit.Parts.Count == 0 && edit.Removed.Count == 0 ? source.Bytes : Repack(package, edit);
    }

    // Every entry in its original order, with its name and time; written parts replace theirs,
    // removed ones are left out, and parts the package did not have go at the end.
    private static byte[] Repack(Package package, PackageEdit edit)
    {
        var parts = new Dictionary<string, byte[]>(edit.Parts, StringComparer.OrdinalIgnoreCase);
        var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in package.Entries)
            {
                var part = entry.FullName.TrimStart('/');
                if (edit.Removed.Contains(part))
                    continue;
                var copy = zip.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                copy.LastWriteTime = entry.LastWriteTime;
                if (entry.FullName.EndsWith('/'))
                    continue;
                using var to = copy.Open();
                if (parts.Remove(part, out var bytes))
                {
                    to.Write(bytes);
                }
                else
                {
                    using var from = package.Open(part);
                    from.CopyTo(to);
                }
            }

            foreach (var (part, bytes) in parts)
            {
                using var to = zip.CreateEntry(part, CompressionLevel.Optimal).Open();
                to.Write(bytes);
            }
        }

        return output.ToArray();
    }
}
