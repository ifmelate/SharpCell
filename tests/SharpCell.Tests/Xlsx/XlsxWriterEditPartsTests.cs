using SharpCell.Xlsx;
using static SharpCell.Tests.Xlsx.XlsxWriterTests;

namespace SharpCell.Tests.Xlsx;

/// <summary>Pictures, comments, charts and pivot table sources follow inserted and deleted rows.</summary>
public class XlsxWriterEditPartsTests
{
    private const string Rels = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static string Anchor(int fromRow, int toRow, string editAs = "") => $"""
        <xdr:twoCellAnchor{editAs}><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{fromRow + 2}" name="Box {fromRow}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>
        """;

    private static TestXlsx Book()
    {
        var rows = string.Concat(Enumerable.Range(1, 8).Select(r => $"<row r=\"{r}\"><c r=\"A{r}\"><v>{r}</v></c></row>"));
        return new TestXlsx().RawSheet("S", $"<sheetData>{rows}</sheetData><drawing r:id=\"rD\"/><legacyDrawing r:id=\"rV\"/>")
            .Part("xl/worksheets/_rels/sheet1.xml.rels", $"""
                <Relationships xmlns="{Rels}">
                <Relationship Id="rD" Type="{R}drawing" Target="../drawings/drawing1.xml"/>
                <Relationship Id="rV" Type="{R}vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
                <Relationship Id="rC" Type="{R}comments" Target="../comments1.xml"/>
                <Relationship Id="rT" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
                </Relationships>
                """)
            .Part("xl/drawings/drawing1.xml", $"<xdr:wsDr xmlns:xdr=\"{Xdr}\" xmlns:a=\"{A}\">{Anchor(4, 6)}{Anchor(2, 3)}{Anchor(6, 7, " editAs=\"absolute\"")}</xdr:wsDr>",
                "application/vnd.openxmlformats-officedocument.drawing+xml")
            .Part("xl/comments1.xml", """
                <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>me</author></authors><commentList>
                <comment ref="B3" authorId="0"><text><t>third</t></text></comment><comment ref="B5" authorId="0"><text><t>fifth</t></text></comment>
                </commentList></comments>
                """, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
            .Part("xl/drawings/vmlDrawing1.vml", """
                <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
                <v:shape id="_x0000_s1025" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Anchor>2, 15, 1, 10, 4, 15, 5, 4</x:Anchor><x:Row>2</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>
                <v:shape id="_x0000_s1026" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Anchor>2, 15, 3, 10, 4, 15, 7, 4</x:Anchor><x:Row>4</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>
                </xml>
                """, "application/vnd.openxmlformats-officedocument.vmlDrawing")
            .Part("xl/threadedComments/threadedComment1.xml", """
                <ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="B5" dT="2026-01-01T00:00:00.00" personId="{00000000-0000-0000-0000-000000000001}" id="{00000000-0000-0000-0000-000000000002}"><text>fifth</text></threadedComment></ThreadedComments>
                """, "application/vnd.ms-excel.threadedcomments+xml")
            .Part("xl/charts/chart1.xml", """
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:val><c:numRef><c:f>S!$A$2:$A$6</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>
                """, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
            .Part("xl/pivotCache/pivotCacheDefinition1.xml", """
                <pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource ref="A1:A8" sheet="S"/></cacheSource><cacheFields count="0"/></pivotCacheDefinition>
                """, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
    }

    private static byte[] Edit(Action<Worksheet> change)
    {
        var original = Book().Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(original));
        workbook.Recalculate();
        change(workbook["S"]);
        workbook.Recalculate();
        var saved = Save(workbook);
        SavedFileValidationTests.AssertNoNewErrors(original, saved);
        return saved;
    }

    [Fact]
    public void Drawings_move_with_the_rows_they_are_anchored_to()
    {
        var drawing = Part(Edit(s => s.InsertRows(3, 2)), "xl/drawings/drawing1.xml");

        Assert.Contains("<xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>", drawing);   // 0-based row 4 (row 5) -> row 7
        Assert.Contains("<xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>", drawing);
        Assert.Contains("<xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>", drawing);   // row 3 is where rows go in: it moves
        Assert.Contains("<xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>", drawing);   // absolute: stays
    }

    [Fact]
    public void A_drawing_in_deleted_rows_shrinks_to_where_they_were()
    {
        var drawing = Part(Edit(s => s.DeleteRows(3, 5)), "xl/drawings/drawing1.xml");

        // Rows 3-7 go: the box on rows 5-7 (0-based 4-6) and the one on rows 3-4 collapse onto row 3.
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(drawing, "<xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>").Count);
    }

    [Fact]
    public void Comments_move_and_those_on_deleted_cells_go()
    {
        var saved = Edit(s => s.DeleteRows(3));

        var comments = Part(saved, "xl/comments1.xml");
        Assert.DoesNotContain("third", comments);
        Assert.Contains("<comment ref=\"B4\"", comments);
        var vml = Part(saved, "xl/drawings/vmlDrawing1.vml");
        Assert.DoesNotContain("_x0000_s1025", vml);
        Assert.Contains("<x:Row>3</x:Row>", vml);
        Assert.Contains("<x:Anchor>2, 15, 2, 10, 4, 15, 6, 4</x:Anchor>", vml);
        Assert.Contains("ref=\"B4\"", Part(saved, "xl/threadedComments/threadedComment1.xml"));
    }

    [Fact]
    public void Chart_series_and_pivot_sources_follow_the_rows()
    {
        var saved = Edit(s => s.InsertRows(3));

        Assert.Contains("<c:f>S!$A$2:$A$7</c:f>", Part(saved, "xl/charts/chart1.xml"));
        Assert.Contains("<worksheetSource ref=\"A1:A9\" sheet=\"S\"", Part(saved, "xl/pivotCache/pivotCacheDefinition1.xml"));
    }
}
