using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SharpCell.Tests.Xlsx;

/// <summary>Locates the reference workbooks under <c>tests/corpus</c>.</summary>
public static class CorpusFiles
{
    public static string Root { get; } = FindRoot();

    /// <summary>Corpus files relative to <see cref="Root"/>, with '/' separators, in a stable order.</summary>
    public static IReadOnlyList<string> All { get; } = Directory.Exists(Root)
        ? Directory.EnumerateFiles(Root, "*.xlsx", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(Root, path).Replace('\\', '/'))
            .Where(path => !Path.GetFileName(path).StartsWith('~'))
            .OrderBy(path => path, System.StringComparer.Ordinal)
            .ToList()
        : [];

    public static string PathOf(string relative) => Path.Combine(Root, relative);

    private static string FindRoot()
    {
        for (var dir = new DirectoryInfo(System.AppContext.BaseDirectory); dir is not null; dir = dir.Parent)
        {
            if (File.Exists(Path.Combine(dir.FullName, "SharpCell.slnx")))
                return Path.Combine(dir.FullName, "tests", "corpus");
        }

        throw new DirectoryNotFoundException("Cannot find the repository root (SharpCell.slnx) above the test binaries.");
    }
}
