using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SharpCell.Conformance;

/// <summary>
/// Runs the reference corpus and writes <c>docs/compatibility.md</c> and <c>docs/compatibility.json</c>.
/// <code>
/// dotnet run --project tools/SharpCell.Conformance [-- options]
///   --corpus DIR    corpus folder (default: tests/corpus)
///   --out DIR       where the report goes (default: docs)
///   --only TEXT     run only files whose path contains TEXT; prints, writes nothing
///   --details       print every mismatching cell
///   --check         write nothing; exit 1 if the committed report differs
/// </code>
/// </summary>
internal static class Program
{
    public static int Main(string[] args)
    {
        var root = RepositoryRoot();
        var corpus = Path.Combine(root, "tests", "corpus");
        var output = Path.Combine(root, "docs");
        string? only = null;
        var details = false;
        var check = false;
        for (var i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--corpus" when i + 1 < args.Length:
                    corpus = Path.GetFullPath(args[++i]);
                    break;
                case "--out" when i + 1 < args.Length:
                    output = Path.GetFullPath(args[++i]);
                    break;
                case "--only" when i + 1 < args.Length:
                    only = args[++i];
                    break;
                case "--details":
                    details = true;
                    break;
                case "--check":
                    check = true;
                    break;
                default:
                    Console.Error.WriteLine($"Unknown option '{args[i]}'.");
                    return 2;
            }
        }

        var results = new List<FileResult>();
        foreach (var file in CorpusFiles(corpus))
        {
            if (only is not null && !file.Contains(only, StringComparison.OrdinalIgnoreCase))
                continue;
            var result = CorpusRunner.Run(Path.Combine(corpus, file), file, TimeSpan.FromMinutes(2));
            results.Add(result);
            if (details || only is not null)
                Print(result, details);
        }

        var report = new Report(results);
        Console.WriteLine($"{report.Passed} of {report.Cells} cells match Excel in {results.Count} files.");
        if (only is not null)
            return 0;

        var json = Path.Combine(output, "compatibility.json");
        var md = Path.Combine(output, "compatibility.md");
        if (check)
        {
            var same = File.Exists(json) && File.ReadAllText(json) == report.ToJson()
                && File.Exists(md) && File.ReadAllText(md) == report.ToMarkdown();
            if (!same)
                Console.Error.WriteLine("The committed compatibility report is out of date; run the tool without --check.");
            return same ? 0 : 1;
        }

        Directory.CreateDirectory(output);
        File.WriteAllText(json, report.ToJson());
        File.WriteAllText(md, report.ToMarkdown());
        Console.WriteLine($"Wrote {md} and {json}.");
        return 0;
    }

    public static IEnumerable<string> CorpusFiles(string corpus) =>
        Directory.EnumerateFiles(corpus, "*.xlsx", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(corpus, path).Replace('\\', '/'))
            .Where(path => !Path.GetFileName(path).StartsWith('~'))
            .OrderBy(path => path, StringComparer.Ordinal);

    private static void Print(FileResult result, bool details)
    {
        Console.WriteLine($"{result.File}: {result.Passed}/{result.Cells.Count}{(result.Error is null ? "" : " " + result.Error)}");
        if (!details)
            return;
        foreach (var cell in result.Cells)
        {
            if (!cell.Passed)
                Console.WriteLine($"  {cell.Sheet}!{cell.Address}  {cell.Formula ?? "(spilled)"}  Excel: {cell.Expected}  SharpCell: {cell.Actual}  [{string.Join(' ', cell.Functions)}]");
        }
    }

    private static string RepositoryRoot()
    {
        for (var dir = new DirectoryInfo(Environment.CurrentDirectory); dir is not null; dir = dir.Parent)
        {
            if (File.Exists(Path.Combine(dir.FullName, "SharpCell.slnx")))
                return dir.FullName;
        }

        for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir is not null; dir = dir.Parent)
        {
            if (File.Exists(Path.Combine(dir.FullName, "SharpCell.slnx")))
                return dir.FullName;
        }

        throw new DirectoryNotFoundException("Run the tool inside the SharpCell repository.");
    }
}
