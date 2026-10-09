using SharpCell;

namespace SharpCell.Tests.Evaluation;

public class ReferenceEvaluationTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s1;
    private readonly Worksheet _data;

    public ReferenceEvaluationTests()
    {
        _s1 = _wb.AddSheet("Sheet1");
        _data = _wb.AddSheet("Data");
        _s1["A1"].Value = 1;
        _s1["B1"].Value = 2;
        _s1["A2"].Value = 3;
        _s1["B2"].Value = "x";
        _data["A1"].Value = 10;
    }

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    private static CellValue Arr(CellValue[,] values) => CellValue.Array(values);

    [Fact]
    public void Cell_references_read_values()
    {
        Assert.Equal(CellValue.Number(3), _wb.Evaluate("=A1+B1"));
        Assert.Equal(CellValue.Number(10), _wb.Evaluate("=Data!A1"));
        Assert.Equal(CellValue.Number(11), _wb.Evaluate("='Data'!A1+A1"));
    }

    [Fact]
    public void Empty_cells_read_as_empty_and_result_in_zero()
    {
        Assert.Equal(CellValue.Number(0), _wb.Evaluate("=Z99"));
        Assert.Equal(CellValue.Text(""), _wb.Evaluate("=Z99&\"\""));
        Assert.Equal(CellValue.True, _wb.Evaluate("=Z99=0"));
        Assert.Equal(CellValue.True, _wb.Evaluate("=Z99=\"\""));
        Assert.Equal(CellValue.True, _wb.Evaluate("=Z99=FALSE"));
    }

    [Fact]
    public void Missing_sheet_is_REF_error()
    {
        Assert.Equal(Err(ErrorKind.Ref), _wb.Evaluate("=Nope!A1"));
        Assert.Equal(Err(ErrorKind.Ref), _wb.Evaluate("=Sheet1!#REF!"));
    }

    [Fact]
    public void Ranges_become_arrays_in_value_context()
    {
        Assert.Equal(Arr(new CellValue[,] { { 1, 2 }, { 3, "x" } }), _wb.Evaluate("=A1:B2"));
        Assert.Equal(Arr(new CellValue[,] { { 10 }, { 30 } }), _wb.Evaluate("=A1:A2*10"));
        Assert.Equal(Arr(new CellValue[,] { { 1, 2, CellValue.Empty } }), _wb.Evaluate("=A1:A1:C1:B1"));
    }

    [Fact]
    public void Range_operator_takes_the_bounding_box()
    {
        Assert.Equal(Arr(new CellValue[,] { { 1, 2 }, { 3, "x" } }), _wb.Evaluate("=B2:A1"));
        Assert.Equal(Arr(new CellValue[,] { { 1, 2 }, { 3, "x" } }), _wb.Evaluate("=(A1):(B2)"));
        Assert.Equal(Err(ErrorKind.Value), _wb.Evaluate("=A1:Data!B2"));
        Assert.Equal(Err(ErrorKind.Value), _wb.Evaluate("=A1:(1)"));
    }

    [Fact]
    public void Intersection_and_union()
    {
        Assert.Equal(CellValue.Text("x"), _wb.Evaluate("=A1:B2 B2:C3"));
        Assert.Equal(Arr(new CellValue[,] { { 2 }, { "x" } }), _wb.Evaluate("=A1:B2 B:B"));
        Assert.Equal(Err(ErrorKind.Null), _wb.Evaluate("=A1:A2 C1:C2"));
        Assert.Equal(Err(ErrorKind.Value), _wb.Evaluate("=(A1,B1)"));
    }

    [Fact]
    public void Three_dimensional_reference_in_value_context_is_VALUE_error()
    {
        Assert.Equal(Err(ErrorKind.Value), _wb.Evaluate("=Sheet1:Data!A1"));
    }

    [Fact]
    public void Defined_names()
    {
        _wb.DefineName("Rate", "=0.2");
        _wb.DefineName("Base", "=Sheet1!$A$2");
        _wb.DefineName("Rate", "=0.5", _data);
        Assert.Equal(CellValue.Number(20), _wb.Evaluate("=100*Rate"));
        Assert.Equal(CellValue.Number(6), _wb.Evaluate("=Base*2"));
        Assert.Equal(CellValue.Number(0.5), _wb.Evaluate("=Data!Rate"));
        Assert.Equal(Err(ErrorKind.Name), _wb.Evaluate("=Unknown+1"));
    }

    [Fact]
    public void Circular_names_are_REF_errors()
    {
        _wb.DefineName("Ping", "=Pong");
        _wb.DefineName("Pong", "=Ping+1");
        Assert.Equal(Err(ErrorKind.Ref), _wb.Evaluate("=Ping"));
    }

    [Fact]
    public void Unsupported_constructs_are_errors_not_exceptions()
    {
        Assert.Equal(Err(ErrorKind.Name), _wb.Evaluate("=NOSUCHFUNCTION(1)"));
    }

    [Fact]
    public void Evaluate_without_sheets_treats_references_as_REF()
    {
        var empty = new Workbook();
        Assert.Equal(Err(ErrorKind.Ref), empty.Evaluate("=A1"));
        Assert.Equal(CellValue.Number(2), empty.Evaluate("=1+1"));
    }
}
