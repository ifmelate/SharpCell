using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using SharpCell;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class XlsxReaderTests
{
    private static CellValue N(double value) => CellValue.Number(value);

    private static Workbook Load(TestXlsx file)
    {
        using var stream = file.Build();
        return XlsxReader.Load(stream);
    }

    [Fact]
    public void Reads_every_kind_of_value()
    {
        var file = new TestXlsx
        {
            SharedStrings = "<si><t>plain</t></si><si><r><t>ri</t></r><r><rPr/><t xml:space=\"preserve\">ch </t></r><rPh><t>ignored</t></rPh></si><si><t>a_x000D_b _x005F_x0041_</t></si>",
        }.Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>1.5</v></c><c r=\"B1\" t=\"s\"><v>0</v></c><c r=\"C1\" t=\"s\"><v>1</v></c><c r=\"D1\" t=\"s\"><v>2</v></c>"
            + "<c r=\"E1\" t=\"b\"><v>1</v></c><c r=\"F1\" t=\"e\"><v>#DIV/0!</v></c><c r=\"G1\" t=\"inlineStr\"><is><t>inline</t></is></c>"
            + "<c r=\"H1\" t=\"str\"><v>formula text</v></c><c r=\"I1\" t=\"d\"><v>2020-01-01T12:00:00</v></c><c r=\"J1\" t=\"e\"><v>#BLOCKED!</v></c>"
            + "<c r=\"K1\" s=\"3\"/><c r=\"L1\"><v>-1E-3</v></c></row>");
        var s = Load(file)["S"];

        Assert.Equal(N(1.5), s["A1"].Value);
        Assert.Equal(CellValue.Text("plain"), s["B1"].Value);
        Assert.Equal(CellValue.Text("rich "), s["C1"].Value);
        Assert.Equal(CellValue.Text("a\rb _x0041_"), s["D1"].Value);
        Assert.Equal(CellValue.True, s["E1"].Value);
        Assert.Equal(CellValue.Error(ErrorKind.Div0), s["F1"].Value);
        Assert.Equal(CellValue.Text("inline"), s["G1"].Value);
        Assert.Equal(CellValue.Text("formula text"), s["H1"].Value);
        Assert.Equal(N(43831.5), s["I1"].Value);
        Assert.Equal(CellValue.Error(ErrorKind.Value), s["J1"].Value);
        Assert.Equal(CellValue.Empty, s["K1"].Value);
        Assert.Equal(N(-0.001), s["L1"].Value);
    }

    [Fact]
    public void Rows_and_cells_without_positions_follow_the_previous_one()
    {
        var file = new TestXlsx().Sheet("S",
            "<row r=\"2\"><c><v>1</v></c><c><v>2</v></c><c r=\"E2\"><v>3</v></c><c><v>4</v></c></row><row><c><v>5</v></c></row>");
        var s = Load(file)["S"];

        Assert.Equal([N(1), N(2), N(3), N(4), N(5)], [s["A2"].Value, s["B2"].Value, s["E2"].Value, s["F2"].Value, s["A3"].Value]);
    }

    [Fact]
    public void Formulas_show_cached_values_until_recalculated()
    {
        var file = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>2</v></c><c r=\"B1\"><f>A1*_xlfn.SINGLE(A1)</f><v>99</v></c><c r=\"C1\" t=\"str\"><f>\"x\"&amp;A1</f><v>x2</v></c></row>");
        var wb = Load(file);
        var s = wb["S"];

        Assert.Equal(N(99), s["B1"].Value);
        Assert.Equal("=A1*@A1", s["B1"].Formula);
        Assert.Equal(CellValue.Text("x2"), s["C1"].Value);

        wb.Recalculate();
        Assert.Equal(N(4), s["B1"].Value);
    }

    [Fact]
    public void Plain_formulas_are_legacy_and_intersect_ranges()
    {
        var file = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"2\"><c r=\"A2\"><v>2</v></c><c r=\"B2\"><f>A1:A3*10</f><v>20</v></c></row><row r=\"3\"><c r=\"A3\"><v>3</v></c></row>");
        var wb = Load(file);
        wb.Recalculate();

        Assert.Equal(N(20), wb["S"]["B2"].Value);
        Assert.Equal(CellValue.Empty, wb["S"]["B3"].Value);
    }

    [Fact]
    public void Shared_formulas_are_translated_to_each_cell()
    {
        var file = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><f t=\"shared\" ref=\"B1:B3\" si=\"0\">A1*2</f><v>2</v></c></row>"
            + "<row r=\"2\"><c r=\"A2\"><v>2</v></c><c r=\"B2\"><f t=\"shared\" si=\"0\"/><v>4</v></c></row>"
            + "<row r=\"3\"><c r=\"A3\"><v>3</v></c><c r=\"B3\"><f t=\"shared\" si=\"0\"/><v>6</v></c></row>");
        var wb = Load(file);
        var s = wb["S"];

        Assert.Equal("=A3*2", s["B3"].Formula);
        s["A3"].Value = 10;
        wb.Recalculate();
        Assert.Equal([N(2), N(4), N(20)], [s["B1"].Value, s["B2"].Value, s["B3"].Value]);
    }

    [Fact]
    public void Shared_formula_without_master_is_unsupported()
    {
        var file = new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"B1\"><f t=\"shared\" si=\"7\"/><v>2</v></c></row>");
        var wb = Load(file);
        wb.Recalculate();

        Assert.Equal(CellValue.Error(ErrorKind.Name), wb["S"]["B1"].Value);
        Assert.Equal(DiagnosticKind.UnsupportedFormula, Assert.Single(wb.Diagnostics).Kind);
    }

    [Fact]
    public void Dynamic_array_formula_takes_back_its_cached_spill()
    {
        var file = new TestXlsx { Metadata = TestXlsx.DynamicArrayMetadata }.Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\" cm=\"1\"><f t=\"array\" ref=\"B1:B2\">A1:A2*10</f><v>10</v></c></row>"
            + "<row r=\"2\"><c r=\"A2\"><v>2</v></c><c r=\"B2\"><v>20</v></c></row>");
        var wb = Load(file);
        var s = wb["S"];
        Assert.Null(s["B2"].Formula);

        s["A2"].Value = 3;
        wb.Recalculate();
        Assert.Equal([N(10), N(30)], [s["B1"].Value, s["B2"].Value]);
    }

    [Fact]
    public void Cm_without_a_metadata_part_counts_as_dynamic()
    {
        var file = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1:A2\">{1;2}</f><v>1</v></c></row><row r=\"2\"><c r=\"A2\"><v>2</v></c></row>");
        var wb = Load(file);
        wb.Recalculate();

        Assert.Equal([N(1), N(2)], [wb["S"]["A1"].Value, wb["S"]["A2"].Value]);
        Assert.Equal(N(3), wb.Evaluate("=SUM(S!A1#)"));
    }

    [Fact]
    public void Array_formula_without_dynamic_flag_fills_its_fixed_area()
    {
        var file = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:C1\">{1,2}</f><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\" t=\"e\"><v>#N/A</v></c></row>");
        var wb = Load(file);
        var s = wb["S"];
        Assert.Throws<InvalidOperationException>(() => s["B1"].Value = 5);

        wb.Recalculate();
        Assert.Equal([N(1), N(2), CellValue.Error(ErrorKind.NA)], [s["A1"].Value, s["B1"].Value, s["C1"].Value]);
    }

    [Fact]
    public void Cm_pointing_at_other_metadata_is_not_dynamic()
    {
        const string metadata = "<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "<metadataTypes count=\"1\"><metadataType name=\"XLRICHVALUE\"/></metadataTypes>"
            + "<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata></metadata>";
        var file = new TestXlsx { Metadata = metadata }.Sheet("S",
            "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1\">{1;2}</f><v>1</v></c></row>");
        var wb = Load(file);
        wb.Recalculate();

        Assert.Equal(N(1), wb["S"]["A1"].Value);
        Assert.Equal(CellValue.Empty, wb["S"]["A2"].Value);
    }

    [Fact]
    public void Reads_the_1904_date_system()
    {
        var wb = Load(new TestXlsx { WorkbookExtra = "<workbookPr date1904=\"1\"/>" }.Sheet("S", ""));

        Assert.Equal(DateSystem.Date1904, wb.DateSystem);
    }

    [Fact]
    public void Defined_names_keep_their_scope_and_skip_built_in_ones()
    {
        var file = new TestXlsx()
            .Sheet("One", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>")
            .ChartSheet("Chart")
            .Sheet("Two", "<row r=\"1\"><c r=\"A1\"><v>2</v></c></row>")
            .Name("Rate", "One!$A$1")
            .Name("Local", "Two!$A$1", " localSheetId=\"2\"")
            .Name("_xlnm.Print_Area", "One!$A$1:$B$2", " localSheetId=\"0\"")
            .Name("_xlnm._FilterDatabase", "One!$A$1", " hidden=\"1\"")
            .Name("Macro", "One!$A$1", " function=\"1\"")
            .Name("OnChart", "One!$A$1", " localSheetId=\"1\"");
        var wb = Load(file);

        Assert.Equal(["One", "Two"], [wb.Sheets[0].Name, wb.Sheets[1].Name]);
        wb["Two"]["B1"].Formula = "=Rate+Local";
        wb["One"]["B1"].Formula = "=Local";
        wb["One"]["C1"].Formula = "=Macro";
        wb.Recalculate();
        Assert.Equal(N(3), wb["Two"]["B1"].Value);
        Assert.Equal(CellValue.Error(ErrorKind.Name), wb["One"]["B1"].Value);
        Assert.Equal(CellValue.Error(ErrorKind.Name), wb["One"]["C1"].Value);
    }

    [Fact]
    public void Links_to_other_workbooks_load_as_unsupported()
    {
        var file = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\"><f>[1]Sheet1!A1+1</f><v>5</v></c><c r=\"B1\"><f>Ext*2</f><v>10</v></c><c r=\"C1\" t=\"str\"><f>\"[1]\"&amp;\"x\"</f><v>[1]x</v></c></row>")
            .Name("Ext", "[1]Sheet1!$A$1");
        var wb = Load(file);
        var s = wb["S"];
        Assert.Equal(N(5), s["A1"].Value);
        Assert.Equal("=[1]Sheet1!A1+1", s["A1"].Formula);

        wb.Recalculate();
        Assert.Equal(CellValue.Error(ErrorKind.Name), s["A1"].Value);
        Assert.Equal(CellValue.Error(ErrorKind.Name), s["B1"].Value);
        Assert.Equal(CellValue.Text("[1]x"), s["C1"].Value);
        Assert.All(wb.Diagnostics, d => Assert.Equal(DiagnosticKind.UnsupportedFormula, d.Kind));
        Assert.Equal(2, wb.Diagnostics.Count);
    }

    [Fact]
    public void Unparsable_formula_keeps_its_text()
    {
        var file = new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><f>SUM(1,</f><v>1</v></c></row>");
        var wb = Load(file);
        wb.Recalculate();

        Assert.Equal("=SUM(1,", wb["S"]["A1"].Formula);
        Assert.Equal(CellValue.Error(ErrorKind.Name), wb["S"]["A1"].Value);
        Assert.Contains("cannot be parsed", Assert.Single(wb.Diagnostics).Message);
    }

    [Fact]
    public void Absolute_relationship_targets_and_hidden_sheets_are_read()
    {
        var file = new TestXlsx { AbsoluteTargets = true, SharedStrings = "<si><t>x</t></si>" }
            .Sheet("Hidden", "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>", " state=\"hidden\"");
        var wb = Load(file);

        Assert.Equal(CellValue.Text("x"), wb["Hidden"]["A1"].Value);
    }

    [Fact]
    public void Data_table_cells_keep_their_cached_values()
    {
        var file = new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><f t=\"dataTable\" ref=\"A1:A2\" dt2D=\"0\" dtr=\"0\" r1=\"B1\"/><v>7</v></c></row>");
        var wb = Load(file);
        wb.Recalculate();

        Assert.Null(wb["S"]["A1"].Formula);
        Assert.Equal(N(7), wb["S"]["A1"].Value);
    }

    [Fact]
    public void Load_from_a_path_and_a_non_seekable_stream()
    {
        var path = Path.Combine(Path.GetTempPath(), $"sharpcell-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var built = new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><v>4</v></c></row>").Build())
            using (var target = File.Create(path))
                built.CopyTo(target);

            Assert.Equal(N(4), XlsxReader.Load(path)["S"]["A1"].Value);

            using var forwardOnly = new ForwardOnlyStream(File.ReadAllBytes(path));
            Assert.Equal(N(4), XlsxReader.Load(forwardOnly)["S"]["A1"].Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Stream_is_left_open()
    {
        using var stream = new TestXlsx().Sheet("S", "").Build();
        XlsxReader.Load(stream);

        Assert.True(stream.CanRead);
    }

    [Fact]
    public void Rejects_legacy_xls_and_encrypted_files()
    {
        using var stream = new MemoryStream([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0]);

        Assert.Throws<NotSupportedException>(() => XlsxReader.Load(stream));
    }

    [Theory]
    [InlineData("application/vnd.ms-excel.sheet.macroEnabled.main+xml")]
    [InlineData("application/vnd.ms-excel.sheet.binary.macroEnabled.main")]
    public void Rejects_macro_enabled_and_binary_workbooks(string contentType)
    {
        var file = new TestXlsx { WorkbookContentType = contentType }.Sheet("S", "");

        Assert.Throws<NotSupportedException>(() => Load(file));
    }

    [Fact]
    public void Rejects_what_is_not_a_zip()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("not a workbook at all"));

        Assert.Throws<InvalidDataException>(() => XlsxReader.Load(stream));
    }

    [Fact]
    public void Rejects_a_package_without_a_workbook()
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
            zip.CreateEntry("readme.txt");
        stream.Position = 0;

        Assert.Throws<InvalidDataException>(() => XlsxReader.Load(stream));
    }

    [Fact]
    public void Rejects_DTDs_and_malformed_xml()
    {
        var withDtd = new TestXlsx().Sheet("S", "").Part("xl/worksheets/sheet1.xml",
            "<?xml version=\"1.0\"?><!DOCTYPE x [<!ENTITY a \"aaaaaaaaaa\">]><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>");
        var broken = new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><v>1</v></row>");

        Assert.Throws<InvalidDataException>(() => Load(withDtd));
        Assert.Throws<InvalidDataException>(() => Load(broken));
    }

    [Theory]
    [InlineData("<row r=\"1\"><c r=\"A1\" t=\"s\"><v>3</v></c></row>")]
    [InlineData("<row r=\"1\"><c r=\"A1\"><v>abc</v></c></row>")]
    [InlineData("<row r=\"1\"><c r=\"A1\" t=\"b\"><v>2</v></c></row>")]
    [InlineData("<row r=\"0\"><c><v>1</v></c></row>")]
    [InlineData("<row r=\"1\"><c r=\"XFE1\"><v>1</v></c></row>")]
    [InlineData("<row r=\"1048577\"><c><v>1</v></c></row>")]
    public void Rejects_invalid_cells(string sheetData)
    {
        Assert.Throws<InvalidDataException>(() => Load(new TestXlsx().Sheet("S", sheetData)));
    }

    [Fact]
    public void Stops_reading_a_part_that_unpacks_past_the_limit()
    {
        var rows = new StringBuilder();
        for (var row = 1; row <= 200; row++)
            rows.Append($"<row r=\"{row}\"><c r=\"A{row}\"><v>{row}</v></c></row>");
        using var stream = new TestXlsx().Sheet("S", rows.ToString()).Build();

        var error = Assert.Throws<InvalidDataException>(() => XlsxReader.Load(stream, new XlsxLimits { MaxPartBytes = 2000 }));
        Assert.Contains("unpacks to more than", error.Message);
    }

    [Fact]
    public void Huge_array_formula_area_does_not_fill_the_sheet()
    {
        var file = new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:XFD1048576\">1</f><v>1</v></c></row>");
        var wb = Load(file);
        RecalculateWithin(wb, TimeSpan.FromSeconds(10));

        Assert.True(wb["S"].Store.Count < 10);
    }

    [Fact]
    public void Huge_cached_spill_area_recalculates_without_allocating_it()
    {
        var file = new TestXlsx().Sheet("S",
            "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1:XFD1048576\">{1;2}</f><v>1</v></c></row><row r=\"2\"><c r=\"A2\"><v>2</v></c></row>");
        var wb = Load(file);
        RecalculateWithin(wb, TimeSpan.FromSeconds(10));

        Assert.Equal([N(1), N(2)], [wb["S"]["A1"].Value, wb["S"]["A2"].Value]);
    }

    [Theory]
    [InlineData("B2:C3")]
    [InlineData("C3:A1")]
    [InlineData("nonsense")]
    public void Array_area_that_does_not_start_at_its_formula_is_one_cell(string reference)
    {
        var file = new TestXlsx().Sheet("S", $"<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"{reference}\">{{1,2}}</f><v>1</v></c></row>");
        var wb = Load(file);
        wb.Recalculate();

        Assert.Equal(N(1), wb["S"]["A1"].Value);
        Assert.Equal(1, wb["S"].Store.Count);
    }

    private static void RecalculateWithin(Workbook workbook, TimeSpan limit)
    {
        using var timeout = new System.Threading.CancellationTokenSource(limit);
        workbook.Recalculate(timeout.Token);
    }

    [Fact]
    public void Array_formulas_share_one_budget_across_the_workbook()
    {
        using var stream = new TestXlsx()
            .Sheet("One", "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:A6\">1</f><v>1</v></c></row>")
            .Sheet("Two", "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:A4\">2</f><v>2</v></c><c r=\"B1\"><f t=\"array\" ref=\"B1:B5\">3</f><v>3</v></c></row>")
            .Build();
        var wb = XlsxReader.Load(stream, new XlsxLimits { MaxSpillCells = 10 });
        wb.Recalculate();

        Assert.Equal([N(1), N(2)], [wb["One"]["A6"].Value, wb["Two"]["A4"].Value]);
        Assert.Equal(CellValue.Error(ErrorKind.Name), wb["Two"]["B1"].Value);
        Assert.Equal(CellValue.Empty, wb["Two"]["B2"].Value);
        Assert.Contains("spill budget", Assert.Single(wb.Diagnostics).Message);
    }

    [Fact]
    public void Array_cells_missing_from_the_file_are_protected_before_calculation()
    {
        var file = new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:A3\">{1;2;3}</f><v>1</v></c></row>");
        var wb = Load(file);

        Assert.Throws<InvalidOperationException>(() => wb["S"]["A3"].Value = 42);
        wb.Recalculate();
        Assert.Equal(N(3), wb["S"]["A3"].Value);
    }

    [Fact]
    public void A_cell_listed_twice_is_invalid()
    {
        var file = new TestXlsx().Sheet("S", "<row r=\"1\"><c r=\"A1\"><f>1+1</f><v>2</v></c><c r=\"A1\"><v>5</v></c></row>");

        Assert.Throws<InvalidDataException>(() => Load(file));
    }

    [Fact]
    public void Newer_errors_saved_as_rich_values_are_read_back()
    {
        const string metadata = "<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xlrd=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\">"
            + "<metadataTypes count=\"1\"><metadataType name=\"XLRICHVALUE\"/></metadataTypes>"
            + "<futureMetadata name=\"XLRICHVALUE\" count=\"2\"><bk><extLst><ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\"><xlrd:rvb i=\"0\"/></ext></extLst></bk>"
            + "<bk><extLst><ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\"><xlrd:rvb i=\"1\"/></ext></extLst></bk></futureMetadata>"
            + "<valueMetadata count=\"2\"><bk><rc t=\"1\" v=\"0\"/></bk><bk><rc t=\"1\" v=\"1\"/></bk></valueMetadata></metadata>";
        var file = new TestXlsx { Metadata = metadata }
            .Sheet("S", "<row r=\"1\"><c r=\"A1\" t=\"e\" vm=\"1\"><v>#VALUE!</v></c><c r=\"B1\" t=\"e\" vm=\"2\"><v>#VALUE!</v></c><c r=\"C1\" t=\"e\"><v>#VALUE!</v></c></row>")
            .Part("xl/richData/rdrichvaluestructure.xml",
                "<rvStructures xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" count=\"1\"><s t=\"_error\"><k n=\"errorType\" t=\"i\"/><k n=\"subType\" t=\"i\"/></s></rvStructures>")
            .Part("xl/richData/rdrichvalue.xml",
                "<rvData xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" count=\"2\"><rv s=\"0\"><v>13</v><v>0</v></rv><rv s=\"0\"><v>8</v><v>0</v></rv></rvData>");
        var s = Load(file)["S"];

        Assert.Equal(CellValue.Error(ErrorKind.Calc), s["A1"].Value);
        Assert.Equal(CellValue.Error(ErrorKind.Spill), s["B1"].Value);
        Assert.Equal(CellValue.Error(ErrorKind.Value), s["C1"].Value);
    }

    [Fact]
    public void Duplicate_sheet_names_are_invalid()
    {
        Assert.Throws<InvalidDataException>(() => Load(new TestXlsx().Sheet("S", "").Sheet("s", "")));
    }

    private sealed class ForwardOnlyStream(byte[] data) : MemoryStream(data)
    {
        public override bool CanSeek => false;
    }
}

public class ExternalReferenceDetectionTests
{
    [Theory]
    [InlineData("[1]Sheet1!A1", true)]
    [InlineData("'[1]My sheet'!A1+1", true)]
    [InlineData("SUM([2]S!A1:A3)", true)]
    [InlineData("[1]!Name", true)]
    [InlineData("Table1[Col]", false)]
    [InlineData("SUM(Table1[[#This Row],[1]])", false)]
    [InlineData("Table1[[1]:[2]]", false)]
    [InlineData("\"[1]\"&A1", false)]
    [InlineData("A1+1", false)]
    public void Recognises_links_to_other_workbooks(string formula, bool expected)
    {
        Assert.Equal(expected, FormulaText.ReferencesOtherWorkbook(formula));
    }

    [Fact]
    public void Zip_directory_that_runs_past_the_end_is_invalid_data()
    {
        // Found by XlsxFuzzTests: one byte raises the extra field length of the first central
        // directory record past the end of the file. .NET 8 reports that as an IOException; the
        // reader promises InvalidDataException for a damaged file.
        var bytes = File.ReadAllBytes(CorpusFiles.PathOf("ironcalc/DynamicArrays.xlsx"));
        Assert.Equal(0x00, bytes[10767]);
        bytes[10767] = 0x2B;
        Assert.Throws<InvalidDataException>(() => XlsxReader.Load(new MemoryStream(bytes)));
    }
}
