using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using SharpCell.Conformance;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>
/// Every corpus file calculated, saved with every value written again, and read back without
/// calculating: the saved results must match Excel wherever the calculated ones did, and nothing
/// but values may change.
/// </summary>
public class RoundTripCorpusTests
{
    internal static readonly XlsxWriteOptions EveryValue = new() { KeepUncalculated = true, RewriteAllValues = true };

    public static TheoryData<string> Files() => [.. CorpusFiles.All];

    [Theory]
    [MemberData(nameof(Files))]
    public void Saved_results_match_Excel_where_calculated_ones_do(string file)
    {
        byte[]? saved = null;
        var result = CorpusRunner.Run(CorpusFiles.PathOf(file), file, TimeSpan.FromMinutes(2), CorpusOverrides.Load(CorpusFiles.Root),
            workbook =>
            {
                var output = new MemoryStream();
                XlsxWriter.Save(workbook, output, EveryValue);
                saved = output.ToArray();
                return XlsxReader.Load(new MemoryStream(saved));
            });

        Assert.True(result.Error is null, $"{file}: {result.Error}");
        var broken = CorpusBaseline.Broken(file, result);
        Assert.True(broken.Count == 0, $"{file}: saved results that no longer match Excel:\n" + string.Join("\n", broken.Take(30)));
        OnlyValuesChanged(file, File.ReadAllBytes(CorpusFiles.PathOf(file)), saved!);
    }

    // Parts a save may write besides worksheets: fullCalcOnLoad and rich errors.
    private static bool MayChange(string name) =>
        name is "xl/workbook.xml" or "xl/metadata.xml" or "[Content_Types].xml" or "xl/_rels/workbook.xml.rels"
        || name.StartsWith("xl/richData/", StringComparison.Ordinal);

    private static void OnlyValuesChanged(string file, byte[] original, byte[] saved)
    {
        using var before = new ZipArchive(new MemoryStream(original));
        using var after = new ZipArchive(new MemoryStream(saved));
        Assert.Equal(before.Entries.Select(e => e.FullName), after.Entries.Take(before.Entries.Count).Select(e => e.FullName));
        Assert.All(after.Entries.Skip(before.Entries.Count), e => Assert.True(MayChange(e.FullName), $"{file}: new part {e.FullName}"));
        foreach (var entry in before.Entries)
        {
            var copy = after.GetEntry(entry.FullName)!;
            if (entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
                Assert.True(XNode.DeepEquals(WithoutCells(entry), WithoutCells(copy)), $"{file}: {entry.FullName} changed outside its cells");
            else if (!MayChange(entry.FullName))
                Assert.True(Bytes(entry).SequenceEqual(Bytes(copy)), $"{file}: {entry.FullName} changed");
        }
    }

    private static XDocument WithoutCells(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        var document = XDocument.Load(stream);
        document.Root!.Elements().Where(e => e.Name.LocalName is "sheetData" or "dimension").Remove();
        return document;
    }

    private static byte[] Bytes(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }
}
