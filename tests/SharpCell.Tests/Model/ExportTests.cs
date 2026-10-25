using System.Data;
using System.Globalization;

namespace SharpCell.Tests.Model;

/// <summary>Ranges read out as a DataTable, as objects and as CSV text, and CSV text read in.</summary>
public class ExportTests
{
    private static (Workbook W, Worksheet S) Book()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sales");
        sheet["A1"].Value = "Region";
        sheet["B1"].Value = "Units";
        sheet["C1"].Value = "Date";
        sheet["D1"].Value = "Mixed";
        sheet["E1"].Value = "Units";
        sheet["A2"].Value = "North";
        sheet["B2"].Value = 1250.5;
        sheet["C2"].Input = "2026-03-05";
        sheet["D2"].Value = 1;
        sheet["E2"].Value = true;
        sheet["A3"].Value = "South, \"far\"";
        sheet["B3"].Formula = "=B2*2";
        sheet["C3"].Input = "2026-04-01";
        sheet["D3"].Formula = "=NA()";
        sheet["B2"].NumberFormat = "#,##0.00";
        return (workbook, sheet);
    }

    [Fact]
    public void A_range_becomes_a_data_table_with_column_types_from_its_values()
    {
        var (_, s) = Book();

        var table = s.Range("A1:F3").ToDataTable();

        Assert.Equal("Sales", table.TableName);
        Assert.Equal(["Region", "Units", "Date", "Mixed", "Units2", "Column6"], table.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
        Assert.Equal([typeof(string), typeof(double), typeof(DateTime), typeof(object), typeof(bool), typeof(object)],
            table.Columns.Cast<DataColumn>().Select(c => c.DataType));
        Assert.Equal(2, table.Rows.Count);
        // B3 was out of date: it is calculated first.
        Assert.Equal(2501.0, table.Rows[1]["Units"]);
        Assert.Equal(new DateTime(2026, 4, 1), table.Rows[1]["Date"]);
        Assert.Equal("#N/A", table.Rows[1]["Mixed"]);
        Assert.Equal(1.0, table.Rows[0]["Mixed"]);
        Assert.Equal(DBNull.Value, table.Rows[1]["Units2"]);
    }

    [Fact]
    public void Without_headers_every_row_is_data()
    {
        var (_, s) = Book();

        var table = s.Range("A2:B3").ToDataTable(headers: false);

        Assert.Equal(["Column1", "Column2"], table.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
        Assert.Equal(2, table.Rows.Count);
    }

    [Fact]
    public void Rows_are_read_into_objects_by_header()
    {
        var (_, s) = Book();

        var rows = s.Range("A:C").ReadRows(row => (Region: row["region"].Text, Units: row["Units"].Value.AsNumber(), Row: row.Row));

        Assert.Equal([("North", 1250.5, 2), ("South, \"far\"", 2501.0, 3)], rows);
        Assert.Throws<KeyNotFoundException>(() => s.Range("A1:C3").ReadRows(row => row["Price"]));
    }

    [Fact]
    public void Csv_text_holds_what_the_cells_show()
    {
        var (_, s) = Book();
        s["A4"].Value = "two\nlines";

        var csv = s.Range("A1:D4").ToCsv();

        Assert.Equal(
            "Region,Units,Date,Mixed\r\n"
            + "North,\"1,250.50\",3/5/2026,1\r\n"
            + "\"South, \"\"far\"\"\",2501,4/1/2026,#N/A\r\n"
            + "\"two\nlines\",,,\r\n", csv);
        Assert.Equal("Region;Units\r\nNorth;1,250.50\r\n", s.Range("A1:B2").ToCsv(';'));
    }

    [Fact]
    public void Csv_text_is_read_as_typed_into_the_cells()
    {
        var workbook = new Workbook { Culture = CultureInfo.GetCultureInfo("en-US") };
        var sheet = workbook.AddSheet("S");
        const string Csv = "Name,Units,Share,Date,Ok\r\n\"Smith, J\",\"1,250\",12%,3/5/2026,TRUE\nLee,=1+1,,'007\r\n\"quoted \"\"word\"\"\nsecond line\",-5";

        var range = sheet.LoadCsv("B2", new StringReader(Csv));
        workbook.Recalculate();

        Assert.Equal("B2:F5", range.Address);
        Assert.Equal("Smith, J", sheet["B3"].Value.AsText());
        Assert.Equal((1250.0, "1,250"), (sheet["C3"].Value.AsNumber(), sheet["C3"].Text));
        Assert.Equal((0.12, "12%"), (sheet["D3"].Value.AsNumber(), sheet["D3"].Text));
        Assert.Equal(new DateTime(2026, 3, 5), sheet["E3"].Value.AsDateTime(DateSystem.Date1900));
        Assert.True(sheet["F3"].Value.AsBoolean());
        // A formula stays text unless asked for.
        Assert.Equal(("=1+1", null), (sheet["C4"].Value.AsText(), sheet["C4"].Formula));
        Assert.Equal(CellValue.Empty, sheet["D4"].Value);
        Assert.Equal("007", sheet["E4"].Value.AsText());
        Assert.Equal("quoted \"word\"\nsecond line", sheet["B5"].Value.AsText());
        Assert.Equal(-5, sheet["C5"].Value.AsNumber());
    }

    [Fact]
    public void Csv_formulas_are_read_when_asked_and_text_round_trips()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");

        sheet.LoadCsv("A1", new StringReader("1;2;=A1+B1"), delimiter: ';', formulas: true);
        workbook.Recalculate();

        Assert.Equal(3, sheet["C1"].Value.AsNumber());
        Assert.Equal("1;2;3\r\n", sheet.Range("A1:C1").ToCsv(';'));
    }
}
