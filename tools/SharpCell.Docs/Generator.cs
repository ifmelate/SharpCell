using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SharpCell.Docs;

/// <summary>Everything the tool writes into the repository, computed in memory, so that check and generate agree.</summary>
internal static class Generator
{
    public const string ApiDirectory = "docs/api";

    public static IReadOnlyList<GeneratedFile> Run(RepositoryLayout layout)
    {
        var snippets = Snippets.Collect(layout.SnippetSources().Select(path => (path, File.ReadAllText(layout.Full(path)))));
        var files = new List<GeneratedFile>();
        foreach (var target in layout.SnippetTargets())
            files.Add(new GeneratedFile(target, Snippets.Apply(File.ReadAllText(layout.Full(target)), snippets, target)));

        var api = ApiReference.Generate(ApiReference.DocumentedAssemblies());
        files.AddRange(api.Files);

        var config = SiteConfig.Load(layout.Full("docs/site.json"));
        var generated = files.ToDictionary(f => f.Path, f => f.Content, StringComparer.Ordinal);
        string Read(string docsPath)
        {
            var path = "docs/" + docsPath;
            return generated.TryGetValue(path, out var content) ? content : File.ReadAllText(layout.Full(path));
        }

        files.Add(new GeneratedFile("docs/llms.txt", LlmsText.Index(config, api.Pages)));
        files.Add(new GeneratedFile("docs/llms-full.txt", LlmsText.Full(config, api.Pages, Read)));
        return files;
    }

    /// <summary>Files that are missing or differ on disk, and stale files left in the generated API folder.</summary>
    public static IReadOnlyList<string> Differences(RepositoryLayout layout, IReadOnlyList<GeneratedFile> files)
    {
        var differences = new List<string>();
        foreach (var file in files)
        {
            var path = layout.Full(file.Path);
            if (!File.Exists(path))
                differences.Add($"{file.Path}: missing");
            else if (File.ReadAllText(path) != file.Content)
                differences.Add($"{file.Path}: out of date");
        }

        foreach (var stale in StaleApiFiles(layout, files))
            differences.Add($"{stale}: not generated any more");
        return differences;
    }

    public static void Write(RepositoryLayout layout, IReadOnlyList<GeneratedFile> files)
    {
        foreach (var file in files)
        {
            var path = layout.Full(file.Path);
            if (File.Exists(path) && File.ReadAllText(path) == file.Content)
                continue;
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, file.Content);
        }

        foreach (var stale in StaleApiFiles(layout, files))
            File.Delete(layout.Full(stale));
    }

    private static IEnumerable<string> StaleApiFiles(RepositoryLayout layout, IReadOnlyList<GeneratedFile> files)
    {
        var api = layout.Full(ApiDirectory);
        if (!Directory.Exists(api))
            return [];
        var owned = files.Select(f => f.Path).ToHashSet(StringComparer.Ordinal);
        return Directory.EnumerateFiles(api, "*", SearchOption.AllDirectories)
            .Select(layout.Relative)
            .Where(path => !owned.Contains(path))
            .Order(StringComparer.Ordinal)
            .ToList();
    }
}
