using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace SharpCell.Xlsx;

/// <summary>A relationship from one part to another; targets are resolved to part names.</summary>
internal sealed record Relationship(string Id, string Type, string Target)
{
    /// <summary>Relationship types differ between transitional and strict files only in their base URI.</summary>
    public bool Is(string kind) => Type.EndsWith("/" + kind, StringComparison.Ordinal);
}

/// <summary>Limits that keep a hostile file from exhausting memory or time.</summary>
internal sealed class XlsxLimits
{
    public static XlsxLimits Default { get; } = new();

    /// <summary>Uncompressed size of one part. A zip bomb is caught when it unpacks past this.</summary>
    public long MaxPartBytes { get; init; } = 1L << 30;

    /// <summary>Uncompressed size of all parts read.</summary>
    public long MaxTotalBytes { get; init; } = 4L << 30;

    /// <summary>
    /// The loaded workbook's spill budget (<see cref="Workbook.MaxSpillCells"/>): cells all spills and
    /// array formulas may cover together, so a tiny file cannot claim billions of cells.
    /// </summary>
    public long MaxSpillCells { get; init; } = 1L << 22;
}

/// <summary>
/// The zip package of an .xlsx file. Part names are compared ignoring case, as the packaging
/// standard requires, and kept without a leading '/'.
/// </summary>
internal sealed class Package : IDisposable
{
    private static readonly byte[] CompoundFileSignature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    private readonly ZipArchive _zip;
    private readonly Dictionary<string, ZipArchiveEntry> _entries = new(StringComparer.OrdinalIgnoreCase);
    private readonly XlsxLimits _limits;
    private long _bytesRead;

    private Package(ZipArchive zip, XlsxLimits limits)
    {
        _zip = zip;
        _limits = limits;
        foreach (var entry in zip.Entries)
            _entries[entry.FullName.TrimStart('/')] = entry;
    }

    public static Package Open(Stream stream, XlsxLimits limits)
    {
        if (!stream.CanSeek)
        {
            var copy = new MemoryStream();
            stream.CopyTo(copy);
            copy.Position = 0;
            stream = copy;
        }

        // Old .xls files and encrypted .xlsx files are both OLE compound files, not zips.
        var start = stream.Position;
        Span<byte> header = stackalloc byte[8];
        var read = stream.ReadAtLeast(header, header.Length, throwOnEndOfStream: false);
        stream.Position = start;
        if (read == header.Length && header.SequenceEqual(CompoundFileSignature))
            throw new NotSupportedException("The file is a legacy .xls workbook or an encrypted workbook; only unencrypted .xlsx files can be read.");

        try
        {
            return new Package(new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true), limits);
        }
        catch (InvalidDataException ex)
        {
            throw new InvalidDataException("The file is not an .xlsx workbook (not a zip package).", ex);
        }
        catch (IOException ex)
        {
            // .NET 8 reports a zip directory that runs past the end of the file as an IOException
            // (later versions as InvalidDataException). The original error stays as the inner one.
            throw new InvalidDataException("The file is not an .xlsx workbook (its zip directory is damaged).", ex);
        }
    }

    public XlsxLimits Limits => _limits;

    public bool Exists(string part) => _entries.ContainsKey(part);

    /// <summary>Opens a part for reading as XML: no DTDs, no external resources, size limited.</summary>
    public XmlReader OpenXml(string part)
    {
        if (!_entries.TryGetValue(part, out var entry))
            throw new InvalidDataException($"The package has no part '{part}'.");

        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            CloseInput = true,
        };
        return XmlReader.Create(new LimitedStream(entry.Open(), this, part), settings);
    }

    /// <summary>The relationships of a part (or of the package for ""), keyed by id. External targets are left out.</summary>
    public Dictionary<string, Relationship> ReadRelationships(string part)
    {
        var slash = part.LastIndexOf('/');
        var folder = slash < 0 ? "" : part[..(slash + 1)];
        var relsPart = folder + "_rels/" + part[(slash + 1)..] + ".rels";
        var result = new Dictionary<string, Relationship>(StringComparer.Ordinal);
        if (!Exists(relsPart))
            return result;

        using var reader = OpenXml(relsPart);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship")
                continue;
            if (reader.GetAttribute("TargetMode") is { } mode && mode.Equals("External", StringComparison.OrdinalIgnoreCase))
                continue;

            var id = reader.GetAttribute("Id");
            var type = reader.GetAttribute("Type");
            var target = reader.GetAttribute("Target");
            if (id is null || type is null || target is null)
                continue;
            result[id] = new Relationship(id, type, ResolveTarget(folder, target));
        }

        return result;
    }

    /// <summary>The content type the package declares for a part, or null.</summary>
    public string? ContentType(string part)
    {
        string? byExtension = null;
        var extension = Path.GetExtension(part).TrimStart('.');
        using var reader = OpenXml("[Content_Types].xml");
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element)
                continue;
            if (reader.LocalName == "Override"
                && string.Equals(reader.GetAttribute("PartName")?.TrimStart('/'), part, StringComparison.OrdinalIgnoreCase))
                return reader.GetAttribute("ContentType");
            if (reader.LocalName == "Default"
                && string.Equals(reader.GetAttribute("Extension"), extension, StringComparison.OrdinalIgnoreCase))
                byExtension = reader.GetAttribute("ContentType");
        }

        return byExtension;
    }

    // Targets are relative to the folder of the source part, or absolute from the package root.
    private static string ResolveTarget(string folder, string target)
    {
        target = Uri.UnescapeDataString(target.Replace('\\', '/'));
        var path = target.StartsWith('/') ? target[1..] : folder + target;
        var segments = new List<string>();
        foreach (var segment in path.Split('/'))
        {
            if (segment is "" or ".")
                continue;
            if (segment == "..")
            {
                if (segments.Count > 0)
                    segments.RemoveAt(segments.Count - 1);
                continue;
            }

            segments.Add(segment);
        }

        return string.Join('/', segments);
    }

    private void Count(int bytes, string part, long partBytes)
    {
        _bytesRead += bytes;
        if (partBytes > _limits.MaxPartBytes)
            throw new InvalidDataException($"Part '{part}' unpacks to more than {_limits.MaxPartBytes} bytes.");
        if (_bytesRead > _limits.MaxTotalBytes)
            throw new InvalidDataException($"The package unpacks to more than {_limits.MaxTotalBytes} bytes.");
    }

    public void Dispose() => _zip.Dispose();

    // Counts what is actually unpacked: sizes declared in the zip directory cannot be trusted.
    private sealed class LimitedStream(Stream inner, Package package, string part) : Stream
    {
        private long _read;

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => _read;
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) => Read(buffer.AsSpan(offset, count));

        public override int Read(Span<byte> buffer)
        {
            var n = inner.Read(buffer);
            _read += n;
            package.Count(n, part, _read);
            return n;
        }

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
