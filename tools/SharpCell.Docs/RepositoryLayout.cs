using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SharpCell.Docs;

/// <summary>Where things are in the repository, as paths relative to its root with '/' separators.</summary>
internal sealed class RepositoryLayout
{
    private static readonly string[] SourceFolders = ["samples", "tests/SharpCell.Tests/Docs"];

    private RepositoryLayout(string root) => Root = root;

    public string Root { get; }

    public static RepositoryLayout At(string root) => new(Path.GetFullPath(root));

    /// <summary>The repository containing the current directory, or else the tool's binaries.</summary>
    public static RepositoryLayout Find()
    {
        foreach (var start in new[] { Environment.CurrentDirectory, AppContext.BaseDirectory })
        {
            for (var dir = new DirectoryInfo(start); dir is not null; dir = dir.Parent)
            {
                if (File.Exists(Path.Combine(dir.FullName, "SharpCell.slnx")))
                    return new RepositoryLayout(dir.FullName);
            }
        }

        throw new DocsException($"Run the tool inside the SharpCell repository: no SharpCell.slnx above {Environment.CurrentDirectory}.");
    }

    public string Full(string relative) => Path.Combine(Root, relative.Replace('/', Path.DirectorySeparatorChar));

    public string Relative(string full) => Path.GetRelativePath(Root, full).Replace('\\', '/');

    /// <summary>The .cs files whose snippets the documentation may show.</summary>
    public IReadOnlyList<string> SnippetSources() =>
        SourceFolders
            .Select(Full)
            .Where(Directory.Exists)
            .SelectMany(dir => Directory.EnumerateFiles(dir, "*.cs", SearchOption.AllDirectories))
            .Select(Relative)
            .Where(path => !path.Contains("/bin/", StringComparison.Ordinal) && !path.Contains("/obj/", StringComparison.Ordinal))
            .Order(StringComparer.Ordinal)
            .ToList();

    /// <summary>The hand-written Markdown files that may contain snippet blocks.</summary>
    public IReadOnlyList<string> SnippetTargets() =>
        new[] { "README.md" }
            .Concat(Directory.EnumerateFiles(Full("docs"), "*.md", SearchOption.AllDirectories).Select(Relative))
            .Where(path => !path.StartsWith(Generator.ApiDirectory + "/", StringComparison.Ordinal) && path != "docs/compatibility.md")
            .Where(path => File.Exists(Full(path)))
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal)
            .ToList();
}
