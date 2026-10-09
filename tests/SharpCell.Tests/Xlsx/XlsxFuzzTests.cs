using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>
/// Corrupted workbooks, at the zip level and inside the XML, may only fail with
/// <see cref="InvalidDataException"/> or <see cref="NotSupportedException"/>; whatever loads must
/// recalculate within a time budget. Seeds are corpus files; mutations are reproducible by seed.
/// </summary>
public class XlsxFuzzTests
{
    // Set SHARPCELL_FUZZ_ITERATIONS for a longer local run.
    private static readonly int Iterations =
        int.TryParse(Environment.GetEnvironmentVariable("SHARPCELL_FUZZ_ITERATIONS"), out var n) ? n : 300;

    private static readonly TimeSpan Budget = TimeSpan.FromSeconds(5);

    private static readonly string[] Seeds =
    [
        "ironcalc/DynamicArrays.xlsx",
        "ironcalc/calc_tests/simple_functions.xlsx",
        "ironcalc/calc_tests/defined_names.xlsx",
        "ironcalc/calc_tests/LOGICAL/IF_ARRAY.xlsx",
        "ironcalc/templates/invoice.xlsx",
        "ironcalc/tables.xlsx",
        "excel-web/filter-mode.xlsx",
    ];

    private static readonly string[] Snippets =
    [
        "<c r=\"A1\">", "<c>", "</c>", "<row>", "<row r=\"0\">", "<row r=\"1048577\">", "<c r=\"XFD1048576\"><v>1</v></c>",
        " t=\"s\"", " t=\"e\"", " t=\"b\"", " t=\"str\"", " t=\"d\"", " t=\"inlineStr\"", " cm=\"1\"", " cm=\"999\"",
        "<v>1e308</v>", "<v>NaN</v>", "<v>-0</v>", "<v>99999999</v>", "<v>#N/A</v>", "<v></v>",
        "<f t=\"shared\" si=\"0\"/>", "<f t=\"shared\" ref=\"A1:A9\" si=\"0\">A1+1</f>", "<f t=\"array\" ref=\"A1:XFD1048576\">1</f>",
        "<f t=\"array\" ref=\"A1:C3\">{1,2}</f>", "<f>SUM(</f>", "<f>A1#</f>", "<f>[1]S!A1</f>", "<f>_xlfn.LAMBDA(_xlpm.x,_xlpm.x)(1)</f>",
        "<f>INDIRECT(\"A\"&amp;ROWS(A1:A9))</f>", "<f>OFFSET(A1,1048575,0)</f>",
        "<definedName name=\"X\">X</definedName>", "<definedName name=\"Y\" localSheetId=\"99\">1</definedName>",
        "<sheet name=\"S\" sheetId=\"9\" r:id=\"rId99\"/>", "<sheet name=\"[bad]\" sheetId=\"9\" r:id=\"rId1\"/>",
        "<workbookPr date1904=\"1\"/>", "<row hidden=\"1\">", "<row r=\"1048576\" hidden=\"1\"/><row hidden=\"1\"/>",
        "<tablePart r:id=\"rId1\"/>", "<tablePart r:id=\"rId99\"/>", "<sheetPr filterMode=\"1\"/>",
        "<autoFilter ref=\"A1:B2\"><filterColumn colId=\"0\"><filters/></filterColumn></autoFilter>",
        "<tableColumn id=\"1\" name=\"X\"/>", " headerRowCount=\"0\"", " totalsRowCount=\"1\"", " ref=\"A1\"", "<!DOCTYPE x>", "&amp;", "&#0;", "_x0000_", "<si><t>", "</si>", "ÿ", "<", ">", "\"",
    ];

    public static TheoryData<string> SeedFiles() => [.. Seeds];

    [Theory]
    [MemberData(nameof(SeedFiles))]
    public void Corrupted_bytes_fail_cleanly(string seed)
    {
        var original = File.ReadAllBytes(CorpusFiles.PathOf(seed));
        var random = new Random(StableSeed(seed));
        for (var i = 0; i < Iterations / 3; i++)
        {
            var bytes = (byte[])original.Clone();
            var flips = random.Next(1, 8);
            for (var f = 0; f < flips; f++)
                bytes[random.Next(bytes.Length)] = (byte)random.Next(256);
            Check(bytes, $"{seed} byte mutation {i}");
        }
    }

    [Theory]
    [MemberData(nameof(SeedFiles))]
    public void Corrupted_xml_fails_cleanly(string seed)
    {
        var parts = Unpack(File.ReadAllBytes(CorpusFiles.PathOf(seed)));
        var names = new List<string>(parts.Keys);
        var random = new Random(StableSeed(seed));
        for (var i = 0; i < Iterations; i++)
        {
            var mutated = new Dictionary<string, string>(parts);
            var edits = random.Next(1, 4);
            for (var e = 0; e < edits; e++)
            {
                var name = names[random.Next(names.Count)];
                mutated[name] = Mutate(mutated[name], random);
            }

            Check(Pack(mutated), $"{seed} xml mutation {i}");
        }
    }

    // string.GetHashCode differs between processes; failures must reproduce.
    private static int StableSeed(string text)
    {
        var hash = 2166136261u;
        foreach (var c in text)
            hash = (hash ^ c) * 16777619u;
        return (int)(hash & 0x7fffffff);
    }

    private static string Mutate(string xml, Random random)
    {
        if (xml.Length == 0)
            return Snippets[random.Next(Snippets.Length)];

        var at = random.Next(xml.Length);
        var length = Math.Min(xml.Length - at, random.Next(1, 40));
        switch (random.Next(4))
        {
            case 0:
                return xml.Remove(at, length);
            case 1:
                return xml.Insert(at, Snippets[random.Next(Snippets.Length)]);
            case 2:
                return xml.Insert(at, xml.Substring(at, length));
            default:
                var chars = xml.ToCharArray();
                chars[at] = "<>\"=/ 0aZ:!#&"[random.Next(13)];
                return new string(chars);
        }
    }

    private static void Check(byte[] bytes, string label)
    {
        var watch = Stopwatch.StartNew();
        try
        {
            using var stream = new MemoryStream(bytes);
            var workbook = XlsxReader.Load(stream);
            using var timeout = new CancellationTokenSource(Budget);
            workbook.Recalculate(timeout.Token);
        }
        catch (InvalidDataException)
        {
        }
        catch (NotSupportedException)
        {
        }
        catch (OperationCanceledException)
        {
            Assert.Fail($"{label}: recalculation took longer than {Budget.TotalSeconds} s.");
        }
        catch (Exception ex)
        {
            Assert.Fail($"{label}: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
        }

        Assert.True(watch.Elapsed < Budget * 2, $"{label}: took {watch.Elapsed.TotalSeconds:0.0} s.");
    }

    private static Dictionary<string, string> Unpack(byte[] bytes)
    {
        var parts = new Dictionary<string, string>(StringComparer.Ordinal);
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        foreach (var entry in zip.Entries)
        {
            using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
            parts[entry.FullName] = reader.ReadToEnd();
        }

        return parts;
    }

    private static byte[] Pack(Dictionary<string, string> parts)
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, content) in parts)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
                writer.Write(content);
            }
        }

        return stream.ToArray();
    }
}
