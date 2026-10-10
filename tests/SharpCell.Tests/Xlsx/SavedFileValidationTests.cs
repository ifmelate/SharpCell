using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>A saved corpus file has no schema errors its source did not have (some Excel files have a few).</summary>
public class SavedFileValidationTests
{
    [Theory]
    [MemberData(nameof(RoundTripCorpusTests.Files), MemberType = typeof(RoundTripCorpusTests))]
    public void A_saved_file_adds_no_schema_errors(string file)
    {
        var original = File.ReadAllBytes(CorpusFiles.PathOf(file));
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        var output = new MemoryStream();
        XlsxWriter.Save(workbook, output, RoundTripCorpusTests.EveryValue);

        var known = Errors(original).ToHashSet();
        var added = Errors(output.ToArray()).Where(e => !known.Contains(e)).ToList();
        Assert.True(added.Count == 0, file + ":\n" + string.Join("\n", added.Take(20)));
    }

    /// <summary>Fails when the saved file has schema errors its source did not have.</summary>
    internal static void AssertNoNewErrors(byte[] original, byte[] saved)
    {
        var known = Errors(original).ToHashSet();
        var added = Errors(saved).Where(e => !known.Contains(e)).ToList();
        Assert.True(added.Count == 0, string.Join("\n", added.Take(20)));
    }

    internal static List<string> Errors(byte[] file)
    {
        using var document = SpreadsheetDocument.Open(new MemoryStream(file, writable: false), isEditable: false);
        return new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document)
            .Select(e => $"{e.Part?.Uri}: {e.Description}").ToList();
    }
}
