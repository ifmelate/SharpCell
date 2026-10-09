using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace SharpCell.Docs;

/// <summary>The compiler's XML documentation, turned into Markdown member by member.</summary>
internal sealed partial class XmlDocs
{
    private readonly Dictionary<string, XElement> _members;

    private XmlDocs(Dictionary<string, XElement> members) => _members = members;

    public static XmlDocs Load(IEnumerable<Assembly> assemblies)
    {
        var members = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var assembly in assemblies)
        {
            var path = Path.ChangeExtension(assembly.Location, ".xml");
            if (!File.Exists(path))
                throw new DocsException($"No XML documentation at {path}; set GenerateDocumentationFile in the library project.");
            Add(members, XDocument.Load(path));
        }

        return new XmlDocs(members);
    }

    public static XmlDocs Parse(string xml)
    {
        var members = new Dictionary<string, XElement>(StringComparer.Ordinal);
        Add(members, XDocument.Parse(xml));
        return new XmlDocs(members);
    }

    public bool Has(string id) => _members.ContainsKey(id);

    public string Summary(string id, Func<string, string?> link) => Block(Effective(id)?.Element("summary"), link);

    public string Remarks(string id, Func<string, string?> link) => Block(Effective(id)?.Element("remarks"), link);

    public string Returns(string id, Func<string, string?> link) => Block(Effective(id)?.Element("returns"), link);

    public IReadOnlyList<(string Name, string Text)> Parameters(string id, Func<string, string?> link) =>
        (Effective(id)?.Elements("param") ?? [])
            .Select(p => ((string?)p.Attribute("name") ?? "", Block(p, link)))
            .ToList();

    public IReadOnlyList<(string Name, string Text)> Exceptions(string id, Func<string, string?> link) =>
        (Effective(id)?.Elements("exception") ?? [])
            .Select(e => (Display((string?)e.Attribute("cref") ?? ""), Block(e, link)))
            .ToList();

    /// <summary>The text up to the first full stop followed by whitespace, or the first paragraph.</summary>
    public static string FirstSentence(string markdown)
    {
        var paragraph = markdown.Split("\n\n")[0];
        var match = SentenceEnd().Match(paragraph);
        return match.Success ? paragraph[..(match.Index + 1)] : paragraph;
    }

    private static void Add(Dictionary<string, XElement> members, XDocument document)
    {
        foreach (var member in document.Descendants("member"))
        {
            if ((string?)member.Attribute("name") is { } name)
                members[name] = member;
        }
    }

    // A member's own tags win; <inheritdoc cref="..."/> supplies the tags it lacks (params by name).
    private XElement? Effective(string id)
    {
        if (!_members.TryGetValue(id, out var member))
            return null;
        var inherit = member.Element("inheritdoc");
        if ((string?)inherit?.Attribute("cref") is not { } cref || Effective(cref) is not { } source)
            return member;

        var merged = new XElement("member", member.Elements().Where(e => e.Name != "inheritdoc"));
        foreach (var element in source.Elements())
        {
            var present = element.Name == "param"
                ? merged.Elements("param").Any(p => (string?)p.Attribute("name") == (string?)element.Attribute("name"))
                : element.Name == "exception"
                    ? merged.Elements("exception").Any(p => (string?)p.Attribute("cref") == (string?)element.Attribute("cref"))
                    : merged.Element(element.Name) is not null;
            if (!present)
                merged.Add(new XElement(element));
        }

        return merged;
    }

    private static string Block(XElement? element, Func<string, string?> link)
    {
        if (element is null)
            return "";
        var raw = string.Concat(element.Nodes().Select(n => Inline(n, link)));
        var paragraphs = raw.Split("\n\n").Select(p => p.Trim()).Where(p => p.Length > 0);
        return string.Join("\n\n", paragraphs);
    }

    private static string Inline(XNode node, Func<string, string?> link) => node switch
    {
        XText text => Whitespace().Replace(text.Value, " "),
        XElement { Name.LocalName: "c" } e => "`" + e.Value.Trim() + "`",
        XElement { Name.LocalName: "see" } e when e.Attribute("langword") is { } word => "`" + word.Value + "`",
        XElement { Name.LocalName: "see" or "seealso" } e when e.Attribute("cref") is { } cref => Link(cref.Value, e.Value.Trim(), link),
        XElement { Name.LocalName: "see" } e when e.Attribute("href") is { } href =>
            $"[{(e.Value.Trim().Length > 0 ? e.Value.Trim() : href.Value)}]({href.Value})",
        XElement { Name.LocalName: "paramref" or "typeparamref" } e => "`" + (string?)e.Attribute("name") + "`",
        XElement { Name.LocalName: "para" } e => "\n\n" + string.Concat(e.Nodes().Select(n => Inline(n, link))).Trim() + "\n\n",
        XElement { Name.LocalName: "code" } e => "\n\n```csharp\n" + Dedent(e.Value) + "\n```\n\n",
        XElement { Name.LocalName: "list" } e => "\n\n" + string.Join("\n", e.Elements("item").Select(item =>
            "- " + string.Concat((item.Element("description") ?? item).Nodes().Select(n => Inline(n, link))).Trim())) + "\n\n",
        XElement { Name.LocalName: "br" } => "\n",
        XElement e => string.Concat(e.Nodes().Select(n => Inline(n, link))),
        _ => "",
    };

    private static string Link(string cref, string text, Func<string, string?> link)
    {
        var label = text.Length > 0 ? text : Display(cref);
        return link(cref) is { } url ? $"[{label}]({url})" : $"`{label}`";
    }

    /// <summary><c>T:N.Type</c> as <c>Type</c>, <c>M:N.Type.Member(...)</c> as <c>Type.Member</c>.</summary>
    private static string Display(string cref)
    {
        var body = cref.Length > 2 && cref[1] == ':' ? cref[2..] : cref;
        var paren = body.IndexOf('(', StringComparison.Ordinal);
        if (paren >= 0)
            body = body[..paren];
        var parts = body.Split('.');
        if (cref.StartsWith("T:", StringComparison.Ordinal) || parts.Length < 2)
            return parts[^1];
        return parts[^1] == "#ctor" ? parts[^2] : parts[^2] + "." + parts[^1];
    }

    private static string Dedent(string code)
    {
        var lines = code.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n').SkipWhile(l => l.Trim().Length == 0).Reverse()
            .SkipWhile(l => l.Trim().Length == 0).Reverse().ToList();
        var indent = lines.Where(l => l.Trim().Length > 0).Select(l => l.Length - l.TrimStart().Length).DefaultIfEmpty(0).Min();
        return string.Join("\n", lines.Select(l => l.Trim().Length == 0 ? "" : l[indent..].TrimEnd()));
    }

    [GeneratedRegex(@"\s+")]
    private static partial Regex Whitespace();

    [GeneratedRegex(@"\.(?=\s)")]
    private static partial Regex SentenceEnd();
}
