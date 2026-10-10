using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

namespace SharpCell.Xlsx;

/// <summary>Writes the file of a <see cref="WritePlan"/>: the source package with the parts that changed.</summary>
internal static class PackageWriter
{
    /// <summary>The file with the workbook's values; the source bytes themselves when nothing changed.</summary>
    public static byte[] Write(WritePlan plan)
    {
        var source = plan.Source;
        using var package = Package.Open(new MemoryStream(source.Bytes, writable: false), source.Limits);
        var relationships = package.ReadRelationships(source.WorkbookPart);
        var saved = WorkbookReader.ReadSavedCells(package, relationships);
        var parts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        foreach (var (sheet, sheetSource) in source.Sheets)
        {
            if (WorksheetWriter.Write(package, sheetSource.Part, sheet, plan, saved) is { } bytes)
                parts[sheetSource.Part] = bytes;
        }

        if (plan.FullCalcOnLoad)
            parts[source.WorkbookPart] = WorkbookPartWriter.WithFullCalcOnLoad(package, source.WorkbookPart);

        return parts.Count == 0 ? source.Bytes : Repack(package, parts);
    }

    // Every entry in its original order, with its name and time; written parts replace theirs, and
    // parts the package did not have go at the end.
    private static byte[] Repack(Package package, Dictionary<string, byte[]> parts)
    {
        var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in package.Entries)
            {
                var copy = zip.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                copy.LastWriteTime = entry.LastWriteTime;
                if (entry.FullName.EndsWith('/'))
                    continue;
                using var to = copy.Open();
                var part = entry.FullName.TrimStart('/');
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
