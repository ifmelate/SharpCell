using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

/// <summary>Notes are read from a file's comments and VML parts and saved back into them.</summary>
public class XlsxNoteTests
{
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    // B2 has a formatted note by Ann, shown; C3 a plain one; D4 a threaded comment, which is no note.
    private static TestXlsx Book() => new TestXlsx()
        .RawSheet("S", """<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><legacyDrawing r:id="rV"/>""")
        .SheetRelationships("S", $"""
            <Relationship Id="rV" Type="{R}vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
            <Relationship Id="rC" Type="{R}comments" Target="../comments1.xml"/>
            <Relationship Id="rT" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
            """)
        .Part("xl/comments1.xml", """
            <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Ann</author><author>tc={1}</author></authors><commentList>
            <comment ref="B2" authorId="0"><text><r><rPr><b/><sz val="9"/><rFont val="Tahoma"/></rPr><t>Ann:</t></r><r><rPr><sz val="9"/><rFont val="Tahoma"/></rPr><t xml:space="preserve">
            check this</t></r></text></comment>
            <comment ref="C3" authorId="0"><text><t>plain</t></text></comment>
            <comment ref="D4" authorId="1"><text><t>[Threaded comment] reply</t></text></comment>
            </commentList></comments>
            """, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
        .Part("xl/drawings/vmlDrawing1.vml", """
            <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
            <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
            <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
            <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;z-index:1"><x:ClientData ObjectType="Note"><x:Anchor>2, 15, 0, 10, 4, 15, 4, 4</x:Anchor><x:Row>1</x:Row><x:Column>1</x:Column><x:Visible/></x:ClientData></v:shape>
            <v:shape id="_x0000_s1026" type="#_x0000_t202" style="position:absolute;z-index:2;visibility:hidden"><x:ClientData ObjectType="Note"><x:Anchor>3, 15, 1, 10, 5, 15, 5, 4</x:Anchor><x:Row>2</x:Row><x:Column>2</x:Column></x:ClientData></v:shape>
            <v:shape id="_x0000_s1027" type="#_x0000_t202" style="position:absolute;z-index:3;visibility:hidden"><x:ClientData ObjectType="Note"><x:Anchor>4, 15, 2, 10, 6, 15, 6, 4</x:Anchor><x:Row>3</x:Row><x:Column>3</x:Column></x:ClientData></v:shape>
            </xml>
            """, "application/vnd.openxmlformats-officedocument.vmlDrawing")
        .Part("xl/threadedComments/threadedComment1.xml", """
            <ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="D4" dT="2026-01-01T00:00:00.00" personId="{00000000-0000-0000-0000-000000000001}" id="{00000000-0000-0000-0000-000000000002}"><text>reply</text></threadedComment></ThreadedComments>
            """, "application/vnd.ms-excel.threadedcomments+xml");

    private static (byte[] Original, Workbook Workbook) Load()
    {
        var original = Book().Build().ToArray();
        return (original, XlsxReader.Load(new MemoryStream(original)));
    }

    private static byte[] SaveChecked(byte[] original, Workbook workbook)
    {
        var saved = Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        Assert.Equal(Presentation.Of(workbook), Presentation.Of(Reload(saved)));
        return saved;
    }

    [Fact]
    public void Notes_are_read_with_their_author_text_and_visibility()
    {
        var (_, workbook) = Load();
        var sheet = workbook["S"];

        Assert.Equal(["B2", "C3"], sheet.Notes.Select(n => n.Cell.Address));
        Assert.Equal(new Note("Ann:\ncheck this", "Ann") { IsVisible = true }, sheet["B2"].Note);
        Assert.Equal(new Note("plain", "Ann"), sheet["C3"].Note);
        Assert.Null(sheet["D4"].Note);
    }

    [Fact]
    public void A_file_whose_notes_did_not_change_is_saved_as_it_was()
    {
        var (original, workbook) = Load();
        workbook["S"]["A1"].Value = 2;

        var saved = SaveChecked(original, workbook);

        Assert.Equal(Part(original.ToArray(), "xl/comments1.xml"), Part(saved, "xl/comments1.xml"));
        Assert.Equal(Part(original.ToArray(), "xl/drawings/vmlDrawing1.vml"), Part(saved, "xl/drawings/vmlDrawing1.vml"));
    }

    [Fact]
    public void Changing_a_note_keeps_the_others_and_the_threaded_comment()
    {
        var (original, workbook) = Load();
        var sheet = workbook["S"];
        sheet["C3"].Note = new Note("changed", "Bob");
        sheet["E5"].Note = new Note("new") { IsVisible = true };

        var saved = SaveChecked(original, workbook);

        var comments = Part(saved, "xl/comments1.xml");
        Assert.Contains("<t>Ann:</t>", comments);                       // B2 kept its runs
        Assert.Contains("[Threaded comment] reply", comments);
        Assert.Contains("<author>Bob</author>", comments);
        var vml = Part(saved, "xl/drawings/vmlDrawing1.vml");
        Assert.Contains("_x0000_s1027", vml);                           // the threaded comment's box
        Assert.Contains("<x:Row>4</x:Row>", vml);                       // E5's new box
        Assert.Contains("_x0000_s1028", vml);
        var reloaded = Reload(saved)["S"];
        Assert.Equal(new Note("changed", "Bob"), reloaded["C3"].Note);
        Assert.True(reloaded["E5"].Note!.IsVisible);
    }

    [Fact]
    public void Removing_a_note_removes_its_box_and_removing_all_keeps_the_threaded_comment()
    {
        var (original, workbook) = Load();
        var sheet = workbook["S"];
        sheet["B2"].Note = null;
        sheet["C3"].Note = new Note("plain", "Ann") { IsVisible = true };

        var saved = SaveChecked(original, workbook);
        var vml = Part(saved, "xl/drawings/vmlDrawing1.vml");
        Assert.DoesNotContain("_x0000_s1025", vml);
        Assert.Contains("<x:Visible", vml);

        sheet["C3"].Note = null;
        saved = SaveChecked(original, workbook);
        Assert.DoesNotContain("plain", Part(saved, "xl/comments1.xml"));
        Assert.Contains("_x0000_s1027", Part(saved, "xl/drawings/vmlDrawing1.vml"));
    }

    [Fact]
    public void Notes_on_a_sheet_without_any_get_their_parts()
    {
        var original = new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>").Sheet("T", "").Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook["T"]["B2"].Note = new Note("on T", "Ann");

        var saved = SaveChecked(original, workbook);

        Assert.Contains("<legacyDrawing r:id=", Part(saved, "xl/worksheets/sheet2.xml").Replace(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"", ""));
        Assert.Contains("o:idmap", Part(saved, "xl/drawings/vmlDrawing1.vml"));
        Assert.Equal("on T", Reload(saved)["T"]["B2"].Note!.Text);
    }

    [Fact]
    public void Removing_every_note_removes_the_parts()
    {
        var original = new TestXlsx().Sheet("S", "").Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook["S"]["A1"].Note = new Note("gone soon");
        var withNote = Save(workbook);
        var reloaded = Reload(withNote);

        reloaded["S"]["A1"].Note = null;
        var saved = SaveChecked(withNote, reloaded);

        Assert.DoesNotContain("legacyDrawing", Part(saved, "xl/worksheets/sheet1.xml"));
        using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(saved));
        Assert.DoesNotContain(zip.Entries, e => e.FullName.EndsWith("comments1.xml", StringComparison.Ordinal) || e.FullName.EndsWith(".vml", StringComparison.Ordinal));
    }

    [Fact]
    public void A_workbook_built_in_code_saves_its_notes()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = "x";
        sheet["A1"].Note = new Note("first\nsecond line", "Ann");
        workbook.AddSheet("T")["C3"].Note = new Note("other sheet");

        var saved = Save(workbook);

        var errors = SavedFileValidationTests.Errors(saved);
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        Assert.Equal(Presentation.Of(workbook), Presentation.Of(Reload(saved)));
        Assert.NotEqual(Part(saved, "xl/drawings/vmlDrawing1.vml").Contains("data=\"1\""), Part(saved, "xl/drawings/vmlDrawing2.vml").Contains("data=\"1\""));
    }
}
