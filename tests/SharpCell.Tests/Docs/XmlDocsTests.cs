using SharpCell.Docs;

namespace SharpCell.Tests.Docs;

public class XmlDocsTests
{
    private const string Xml = """
        <?xml version="1.0"?>
        <doc><members>
          <member name="T:N.A">
            <summary>
              Reads <c>x</c>   and   writes <see cref="T:N.B"/>, see <see cref="M:N.B.Run(System.String)">Run</see>,
              <see langword="null"/>, <see cref="T:System.IO.Stream"/> and <paramref name="path"/>.
              <para>Second paragraph.</para>
            </summary>
            <remarks><list type="bullet"><item><description>one</description></item><item><description>two</description></item></list></remarks>
          </member>
          <member name="M:N.A.Load(System.String)">
            <summary>Loads it.</summary>
            <param name="path">The file.</param>
            <returns>The thing.</returns>
            <exception cref="T:System.IO.IOException">The file cannot be read.</exception>
          </member>
          <member name="M:N.A.Load(System.IO.Stream)">
            <summary>Loads from a stream.</summary>
            <inheritdoc cref="M:N.A.Load(System.String)"/>
          </member>
        </members></doc>
        """;

    private static string? Link(string cref) => cref switch
    {
        "T:N.B" => "N.B.md",
        "M:N.B.Run(System.String)" => "N.B.md#run-string",
        _ => null,
    };

    [Fact]
    public void Summary_becomes_markdown_with_links_for_documented_types_only()
    {
        var docs = XmlDocs.Parse(Xml);
        Assert.Equal(
            "Reads `x` and writes [B](N.B.md), see [Run](N.B.md#run-string), `null`, `Stream` and `path`.\n\nSecond paragraph.",
            docs.Summary("T:N.A", Link));
        Assert.Equal("- one\n- two", docs.Remarks("T:N.A", Link));
    }

    [Fact]
    public void Inheritdoc_fills_what_the_member_does_not_say_itself()
    {
        var docs = XmlDocs.Parse(Xml);
        Assert.Equal("Loads from a stream.", docs.Summary("M:N.A.Load(System.IO.Stream)", Link));
        Assert.Equal("The thing.", docs.Returns("M:N.A.Load(System.IO.Stream)", Link));
        Assert.Equal(new[] { ("path", "The file.") }, docs.Parameters("M:N.A.Load(System.IO.Stream)", Link));
        Assert.Equal(new[] { ("IOException", "The file cannot be read.") }, docs.Exceptions("M:N.A.Load(System.IO.Stream)", Link));
    }

    [Fact]
    public void First_sentence_stops_at_the_first_full_stop_followed_by_a_space()
    {
        Assert.Equal("Reads `x.y` and writes.", XmlDocs.FirstSentence("Reads `x.y` and writes. Then more."));
        Assert.Equal("No full stop", XmlDocs.FirstSentence("No full stop"));
        Assert.Equal("One line.", XmlDocs.FirstSentence("One line.\n\nSecond paragraph."));
    }
}
