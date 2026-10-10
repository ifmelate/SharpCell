using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace SharpCell.Xlsx;

/// <summary>Copying and saving XML parts without changing what is not ours to change.</summary>
internal static class XmlParts
{
    /// <summary>UTF-8 without a byte order mark; the declaration is copied from the source as it is.</summary>
    public static XmlWriterSettings WriterSettings { get; } = new()
    {
        Encoding = new UTF8Encoding(false),
        OmitXmlDeclaration = true,
        CloseOutput = false,
    };

    /// <summary>Writes the start tag of the element the reader is on, with its attributes, and moves into it.</summary>
    public static void CopyStartTag(XmlReader reader, XmlWriter writer)
    {
        writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
        writer.WriteAttributes(reader, defattr: true);
        reader.MoveToElement();
        reader.Read();
    }

    /// <summary>Copies the node the reader is on (an element with everything inside it) and moves past it.</summary>
    public static void CopyNode(XmlReader reader, XmlWriter writer)
    {
        switch (reader.NodeType)
        {
            case XmlNodeType.Element:
                writer.WriteNode(reader, defattr: true);
                return;
            case XmlNodeType.XmlDeclaration:
                writer.WriteProcessingInstruction(reader.Name, reader.Value);
                break;
            case XmlNodeType.Text:
                writer.WriteString(reader.Value);
                break;
            case XmlNodeType.CDATA:
                writer.WriteCData(reader.Value);
                break;
            case XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace:
                writer.WriteWhitespace(reader.Value);
                break;
            case XmlNodeType.EndElement:
                writer.WriteFullEndElement();
                break;
        }

        reader.Read();
    }

    /// <summary>A small part loaded as a document, saved with its declaration, UTF-8 without a byte order mark.</summary>
    public static byte[] Save(XDocument document)
    {
        var output = new MemoryStream();
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false) }))
            document.Save(writer);
        return output.ToArray();
    }
}
