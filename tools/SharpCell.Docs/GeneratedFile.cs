namespace SharpCell.Docs;

/// <summary>A file the tool owns: its path from the repository root, with '/' separators, and its content.</summary>
internal sealed record GeneratedFile(string Path, string Content);
