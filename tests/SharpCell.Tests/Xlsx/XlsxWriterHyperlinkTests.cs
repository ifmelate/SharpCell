using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

/// <summary>Hyperlinks are saved: the address in the sheet's relationships, the place in the workbook in location.</summary>
public class XlsxWriterHyperlinkTests
{
    private const string HyperlinkType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    private const string Sheet1 = "xl/worksheets/sheet1.xml";
    private const string Sheet1Relationships = "xl/worksheets/_rels/sheet1.xml.rels";

    private static string External(string id, string target) =>
        $"<Relationship Id=\"{id}\" Type=\"{HyperlinkType}\" Target=\"{target}\" TargetMode=\"External\"/>";

    private static TestXlsx Linked() => new TestXlsx().RawSheet("S", """
        <sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="5"><c r="A5"><v>5</v></c></row></sheetData>
        <hyperlinks><hyperlink ref="A1" r:id="rId1" tooltip="Home"/><hyperlink ref="B2" location="Other!A1" display="Other"/><hyperlink ref="C3:C4" r:id="rId2"/></hyperlinks>
        <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        """).Sheet("Other", "").SheetRelationships("S", External("rId1", "https://example.com/") + External("rId2", "mailto:a@example.com"));

    private static TestXlsx Unlinked() => new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>").Sheet("Other", "");

    public static TheoryData<string, string> Changes => new()
    {
        { "linked", "external added" }, { "linked", "in workbook added" }, { "linked", "address changed" },
        { "linked", "tooltip changed" }, { "linked", "one removed" }, { "linked", "all removed" },
        { "linked", "rows inserted" }, { "linked", "rows deleted" }, { "linked", "columns deleted" },
        { "unlinked", "external added" }, { "unlinked", "in workbook added" },
    };

    private static void Change(Worksheet s, string change)
    {
        switch (change)
        {
            case "external added": s["D7"].Hyperlink = new Hyperlink("https://example.com/new?a=1&b=2", "A <new> \"one\""); break;
            case "in workbook added": s["E1"].Hyperlink = new Hyperlink("#Other!B5"); break;
            case "address changed": s["A1"].Hyperlink = new Hyperlink("https://example.org/"); break;
            case "tooltip changed": s["A1"].Hyperlink = new Hyperlink("https://example.com/", "Away"); break;
            case "one removed": s["C3"].Hyperlink = null; break;
            case "all removed":
                foreach (var cell in new[] { "A1", "B2", "C3", "C4" })
                    s[cell].Hyperlink = null;
                break;
            case "rows inserted": s.InsertRows(2, 3); break;
            case "rows deleted": s.DeleteRows(1, 2); break;
            case "columns deleted": s.DeleteColumns(2); break;
        }
    }

    // Each cell's link, in row then column order.
    private static string Links(Worksheet sheet) => string.Join("\n", sheet.Hyperlinks
        .OrderBy(link => link.Key.Row).ThenBy(link => link.Key.Column)
        .Select(link => $"{link.Key} {link.Value.Target} [{link.Value.Tooltip}]"));

    [Theory]
    [MemberData(nameof(Changes))]
    public void A_change_is_saved_and_reads_back(string file, string change)
    {
        var original = (file == "linked" ? Linked() : Unlinked()).Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        Change(workbook["S"], change);
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        var reloaded = XlsxWriterTests.Reload(saved);
        Assert.Equal(Links(workbook["S"]), Links(reloaded["S"]));
        Assert.Equal(workbook["S"]["A1"].Value, reloaded["S"]["A1"].Value);
    }

    [Fact]
    public void A_workbook_built_in_code_saves_its_links()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Report");
        workbook.AddSheet("Sheet2");
        sheet["A1"].Value = "Site";
        sheet["A1"].Hyperlink = new Hyperlink("https://example.com/", "Our site");
        sheet["A2"].Hyperlink = new Hyperlink("mailto:sales@example.com?subject=Order");
        sheet["A3"].Hyperlink = new Hyperlink("#Sheet2!B2");
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        var reloaded = XlsxWriterTests.Reload(saved);
        Assert.Equal(Links(sheet), Links(reloaded["Report"]));
        Assert.Equal("Site", reloaded["Report"]["A1"].Value.AsText());
    }

    [Fact]
    public void Changing_a_link_rewrites_only_the_links_and_their_relationships()
    {
        var original = Linked().Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook["S"]["A1"].Hyperlink = new Hyperlink("https://example.org/");
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        Assert.Equal([Sheet1Relationships, Sheet1], ChangedParts(original, saved));
        Assert.True(XNode.DeepEquals(WithoutLinks(original), WithoutLinks(saved)));
        var relationships = XlsxWriterTests.Part(saved, Sheet1Relationships);
        Assert.Contains("Target=\"https://example.org/\"", relationships);
        Assert.DoesNotContain("https://example.com/", relationships);
        Assert.Contains("Target=\"mailto:a@example.com\"", relationships);
    }

    [Fact]
    public void Removing_every_link_removes_the_element_and_the_relationships()
    {
        var workbook = XlsxWriterTests.Load(Linked());
        foreach (var cell in new[] { "A1", "B2", "C3", "C4" })
            workbook["S"][cell].Hyperlink = null;

        var saved = XlsxWriterTests.Save(workbook);

        Assert.DoesNotContain("hyperlink", XlsxWriterTests.Part(saved, Sheet1));
        Assert.DoesNotContain("hyperlink", XlsxWriterTests.Part(saved, Sheet1Relationships));
    }

    [Fact]
    public void A_workbook_with_links_saved_without_changes_is_the_source_bytes()
    {
        var original = Linked().RawSheet("Wide", """
            <sheetData/><hyperlinks><hyperlink ref="A:A" r:id="rIdK"/><hyperlink ref="B1" location="S!A1"/></hyperlinks>
            """).SheetRelationships("Wide", External("rIdK", "https://kept.example.com/")).Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();

        Assert.Equal(original, XlsxWriterTests.Save(workbook));
    }

    [Fact]
    public void Links_follow_inserted_rows_when_saved()
    {
        var workbook = XlsxWriterTests.Load(Linked());
        workbook["S"].InsertRows(2, 2);
        workbook.Recalculate();

        var reloaded = XlsxWriterTests.Reload(XlsxWriterTests.Save(workbook));

        Assert.Equal("""
            A1 https://example.com/ [Home]
            B4 #Other!A1 []
            C5 mailto:a@example.com []
            C6 mailto:a@example.com []
            """.ReplaceLineEndings("\n").TrimEnd('\n'), Links(reloaded["S"]));
    }

    [Fact]
    public void A_link_inserted_into_a_linked_range_does_not_spread()
    {
        var workbook = XlsxWriterTests.Load(Linked());
        workbook["S"].InsertRows(4);   // between C3 and C4
        workbook.Recalculate();

        var reloaded = XlsxWriterTests.Reload(XlsxWriterTests.Save(workbook));

        Assert.Equal(Links(workbook["S"]), Links(reloaded["S"]));
        Assert.Null(reloaded["S"]["C4"].Hyperlink);
    }

    [Fact]
    public void Fragments_and_quoted_sheet_names_survive_a_round_trip()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Index");
        workbook.AddSheet("Q1 Sales");
        workbook.DefineName("MyName", "='Q1 Sales'!$B$2");
        sheet["A1"].Hyperlink = new Hyperlink("https://example.com/page#part");
        sheet["A2"].Hyperlink = new Hyperlink("#'Q1 Sales'!A1", "Sales");
        sheet["A3"].Hyperlink = new Hyperlink("#Index!C3");
        sheet["A4"].Hyperlink = new Hyperlink("#MyName");
        workbook.Recalculate();

        var saved = XlsxWriterTests.Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        var part = XlsxWriterTests.Part(saved, Sheet1);
        var relationships = XlsxWriterTests.Part(saved, Sheet1Relationships);
        Assert.Matches("<hyperlink ref=\"A1\" r:id=\"[^\"]+\" location=\"part\" */>", part);
        Assert.Contains("<hyperlink ref=\"A2\" location=\"'Q1 Sales'!A1\" tooltip=\"Sales\"", part);
        Assert.Contains("<hyperlink ref=\"A3\" location=\"Index!C3\"", part);
        Assert.Contains("<hyperlink ref=\"A4\" location=\"MyName\"", part);
        Assert.Contains($"Type=\"{HyperlinkType}\" Target=\"https://example.com/page\" TargetMode=\"External\"", relationships);
        Assert.Single(Regex.Matches(relationships, "<Relationship "));
        Assert.Equal(Links(sheet), Links(XlsxWriterTests.Reload(saved)["Index"]));
    }

    [Fact]
    public void Hyperlink_relationships_are_shared_and_removed()
    {
        var file = new TestXlsx().RawSheet("S", """
            <sheetData/><hyperlinks><hyperlink ref="A:A" r:id="rIdK"/><hyperlink ref="D1" r:id="rId1"/><hyperlink ref="D2" r:id="rId2"/></hyperlinks>
            """).SheetRelationships("S", External("rIdK", "https://kept.example.com/")
            + External("rId1", "https://example.com/d") + External("rId2", "https://example.com/d"));
        var workbook = XlsxWriterTests.Load(file);
        var sheet = workbook["S"];
        sheet["B2"].Hyperlink = new Hyperlink("https://example.com/x");
        sheet["B3"].Hyperlink = new Hyperlink("https://example.com/x", "Same place");

        var shared = XlsxWriterTests.Save(workbook);

        var relationships = XlsxWriterTests.Part(shared, Sheet1Relationships);
        Assert.Single(Regex.Matches(relationships, "Target=\"https://example.com/x\""));
        Assert.Single(Regex.Matches(relationships, "Target=\"https://example.com/d\""));
        Assert.Contains("Id=\"rIdK\"", relationships);
        var part = XlsxWriterTests.Part(shared, Sheet1);
        Assert.Contains("<hyperlink ref=\"A:A\" r:id=\"rIdK\"", part);
        var ids = Regex.Matches(part, "<hyperlink ref=\"B[23]\" r:id=\"([^\"]+)\"").Select(m => m.Groups[1].Value).ToList();
        Assert.Equal(2, ids.Count);
        Assert.Equal(ids[0], ids[1]);
        SavedFileValidationTests.AssertNoNewErrors(file.Build().ToArray(), shared);

        var again = XlsxWriterTests.Reload(shared);
        again.Recalculate();
        foreach (var cell in new[] { "B2", "B3", "D1", "D2" })
            again["S"][cell].Hyperlink = null;
        var removed = XlsxWriterTests.Save(again);

        relationships = XlsxWriterTests.Part(removed, Sheet1Relationships);
        Assert.DoesNotContain("example.com/x", relationships);
        Assert.DoesNotContain("example.com/d", relationships);
        Assert.Contains("Id=\"rIdK\" Type=\"" + HyperlinkType + "\" Target=\"https://kept.example.com/\"", relationships);
        part = XlsxWriterTests.Part(removed, Sheet1);
        Assert.Contains("<hyperlink ref=\"A:A\" r:id=\"rIdK\"", part);
        Assert.DoesNotContain("ref=\"B2\"", part);
    }

    // Entries whose bytes differ between the files, or that only one of them has.
    internal static List<string> ChangedParts(byte[] a, byte[] b)
    {
        var first = Entries(a);
        var second = Entries(b);
        return [.. first.Keys.Union(second.Keys).Order(StringComparer.Ordinal)
            .Where(name => !first.TryGetValue(name, out var x) || !second.TryGetValue(name, out var y) || !x.SequenceEqual(y))];
    }

    private static Dictionary<string, byte[]> Entries(byte[] file)
    {
        using var zip = new ZipArchive(new MemoryStream(file));
        return zip.Entries.ToDictionary(entry => entry.FullName, entry => XlsxWriterTests.PartBytes(file, entry.FullName));
    }

    private static XDocument WithoutLinks(byte[] file)
    {
        var document = XDocument.Parse(XlsxWriterTests.Part(file, Sheet1));
        document.Root!.Elements().Where(e => e.Name.LocalName == "hyperlinks").Remove();
        return document;
    }
}
