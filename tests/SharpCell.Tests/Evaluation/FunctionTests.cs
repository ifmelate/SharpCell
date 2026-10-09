using SharpCell;
using SharpCell.Evaluation;
using SharpCell.Functions;
using SharpCell.Parsing;

namespace SharpCell.Tests.Evaluation;

public class FunctionTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public FunctionTests()
    {
        _s = _wb.AddSheet("Sheet1");
        var s2 = _wb.AddSheet("Sheet2");
        _s["A1"].Value = 1;
        _s["A2"].Value = "x";
        _s["A3"].Value = true;
        _s["B1"].Value = 10;
        _s["B2"].Value = 20;
        _s["C1"].Value = CellValue.Error(ErrorKind.NA);
        s2["A1"].Value = 5;
        for (var row = 1; row <= 50; row++)
            _s[row * 1000, 4].Value = row;
    }

    private static CellValue Err(ErrorKind kind) => CellValue.Error(kind);

    private static CellValue Arr(CellValue[,] values) => CellValue.Array(values);

    public static TheoryData<string, CellValue> Aggregates => new()
    {
        { "=SUM(1,2,3)", 6 },
        { "=SUM(A1:A3)", 1 },
        { "=SUM(\"2\",TRUE)", 3 },
        { "=SUM(\"x\")", Err(ErrorKind.Value) },
        { "=SUM(A1:A3,#N/A)", Err(ErrorKind.NA) },
        { "=SUM(A1:C1)", Err(ErrorKind.NA) },
        { "=SUM({1,\"2\",TRUE})", 1 },
        { "=SUM(D:D)", 1275 },
        { "=SUM(Sheet1:Sheet2!A1)", 6 },
        { "=SUM((A1,B1:B2))", 31 },
        { "=SUM(1,)", 1 },
        { "=AVERAGE(1,2)", 1.5 },
        { "=AVERAGE(B1:B2,A1:A3)", 31.0 / 3 },
        { "=AVERAGE(E1:E5)", Err(ErrorKind.Div0) },
        { "=MIN(3,1,2)", 1 },
        { "=MAX(A1:A3,B1:B2)", 20 },
        { "=MAX(A2:A3)", 0 },
        { "=COUNT(1,\"2\",\"x\",TRUE,#N/A)", 3 },
        { "=COUNT(A1:C1,B1:B2)", 4 },
        { "=COUNTA(1,\"\",#N/A)", 3 },
        { "=COUNTA(A1:C3)", 6 },
    };

    [Theory]
    [MemberData(nameof(Aggregates))]
    public void Aggregate_functions(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    public static TheoryData<string, CellValue> Logical => new()
    {
        { "=IF(TRUE,1,2)", 1 },
        { "=IF(0,1,2)", 2 },
        { "=IF(FALSE,1)", false },
        { "=IF(FALSE,1,)", 0 },
        { "=IF(FALSE,1,)&\"x\"", "0x" },
        { "=IF(\"x\",1,2)", Err(ErrorKind.Value) },
        { "=IF(#N/A,1,2)", Err(ErrorKind.NA) },
        { "=IF(TRUE,1,1/0)", 1 },
        { "=SUM(IF(TRUE,B1:B2,0))", 30 },
        { "=IF({TRUE,FALSE},{1,2},{3,4})", Arr(new CellValue[,] { { 1, 4 } }) },
        { "=IF({1,0},\"a\",\"b\")", Arr(new CellValue[,] { { "a", "b" } }) },
        { "=IFERROR(1/0,\"x\")", "x" },
        { "=IFERROR(5,1/0)", 5 },
        { "=IFERROR(1/0,)", 0 },
        { "=IFERROR({1,#N/A},0)", Arr(new CellValue[,] { { 1, 0 } }) },
        { "=CHOOSE(2,\"a\",\"b\")", "b" },
        { "=CHOOSE(1.9,\"a\",\"b\")", "a" },
        { "=CHOOSE(3,\"a\",\"b\")", Err(ErrorKind.Value) },
        { "=CHOOSE(0,\"a\",\"b\")", Err(ErrorKind.Value) },
        { "=SUM(CHOOSE(2,A1:A2,B1:B2))", 30 },
        { "=CHOOSE({1,2},\"a\",\"b\")", Arr(new CellValue[,] { { "a", "b" } }) },
        { "=AND(TRUE,1)", true },
        { "=AND(TRUE,0)", false },
        { "=OR(FALSE,0)", false },
        { "=OR(FALSE,2)", true },
        { "=AND(A1:A3)", true },
        { "=AND(\"x\")", Err(ErrorKind.Value) },
        { "=AND(\"TRUE\")", true },
        { "=AND(E1:E5)", Err(ErrorKind.Value) },
        { "=OR(A1:C1)", Err(ErrorKind.NA) },
        { "=NOT(0)", true },
        { "=NOT(\"x\")", Err(ErrorKind.Value) },
        { "=NOT({TRUE,FALSE})", Arr(new CellValue[,] { { false, true } }) },
    };

    [Theory]
    [MemberData(nameof(Logical))]
    public void Logical_functions(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    public static TheoryData<string, CellValue> Information => new()
    {
        { "=ABS(-2)", 2 },
        { "=ABS(\"-3\")", 3 },
        { "=ABS(B1:B2*-1)", Arr(new CellValue[,] { { 10 }, { 20 } }) },
        { "=ISBLANK(Z1)", true },
        { "=ISBLANK(\"\")", false },
        { "=ISBLANK(A1:A4)", Arr(new CellValue[,] { { false }, { false }, { false }, { true } }) },
        { "=ISERROR(1/0)", true },
        { "=ISERROR(1)", false },
        { "=ROWS(A1:B3)", 3 },
        { "=COLUMNS(A1:B3)", 2 },
        { "=ROWS({1;2;3})", 3 },
        { "=COLUMNS(5)", 1 },
        { "=ROWS(A:A)", 1048576 },
        { "=ROWS((A1,B1))", Err(ErrorKind.Ref) },
        { "=ABS()", Err(ErrorKind.Value) },
        { "=ABS(1,2)", Err(ErrorKind.Value) },
        { "=sum(1,2)", 3 },
        { "=_xlfn.SUM(1)", 1 },
    };

    [Theory]
    [MemberData(nameof(Information))]
    public void Other_functions(string formula, CellValue expected) => Assert.Equal(expected, _wb.Evaluate(formula));

    [Fact]
    public void Untaken_branch_is_not_evaluated()
    {
        var context = new EvaluationContext(_wb, _s, new CellAddress(1, 1));
        Evaluator.EvaluateFormula(FormulaParser.Parse("=IF(TRUE,1,RAND())", new CellAddress(1, 1)), context);
        Assert.False(context.UsedVolatile);
        Evaluator.EvaluateFormula(FormulaParser.Parse("=IF(FALSE,1,RAND())", new CellAddress(1, 1)), context);
        Assert.True(context.UsedVolatile);
    }

    private sealed class FixedClock(DateTime local) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => new(local, TimeSpan.Zero);

        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }

    [Fact]
    public void NOW_uses_the_workbook_clock_and_date_system()
    {
        var now = new DateTime(2026, 10, 8, 12, 0, 0);
        _wb.Clock = new FixedClock(now);
        var serial = (now - new DateTime(1899, 12, 30)).TotalDays;
        Assert.Equal(CellValue.Number(serial), _wb.Evaluate("=NOW()"));
        _wb.DateSystem = DateSystem.Date1904;
        Assert.Equal(CellValue.Number(serial - 1462), _wb.Evaluate("=NOW()"));
    }

    [Fact]
    public void RAND_uses_the_workbook_random_source()
    {
        _wb.Random = new Random(42);
        var expected = new Random(42).NextDouble();
        Assert.Equal(CellValue.Number(expected), _wb.Evaluate("=RAND()"));
    }

    [Fact]
    public void Exception_in_a_function_becomes_VALUE_and_a_diagnostic()
    {
        var registry = new FunctionRegistry();
        registry.Add(new FunctionInfo("BOOM", 0, 0, [ArgumentKind.Value], _ => throw new InvalidOperationException("kaboom")));
        _wb.Functions = registry;

        Assert.Equal(Err(ErrorKind.Value), _wb.Evaluate("=BOOM()"));
        var diagnostic = Assert.Single(_wb.Diagnostics);
        Assert.Equal(DiagnosticKind.FunctionFailure, diagnostic.Kind);
        Assert.Contains("BOOM", diagnostic.Message);
        Assert.Contains("kaboom", diagnostic.Message);
    }

    [Fact]
    public void Cancellation_is_not_swallowed()
    {
        var registry = new FunctionRegistry();
        registry.Add(new FunctionInfo("STOP", 0, 0, [ArgumentKind.Value], _ => throw new OperationCanceledException()));
        _wb.Functions = registry;
        Assert.Throws<OperationCanceledException>(() => _wb.Evaluate("=STOP()"));
    }

    [Fact]
    public void Not_implemented_functions_are_NAME_errors()
    {
        var registry = new FunctionRegistry();
        registry.Add(new FunctionInfo("LATER", 0, 0, [ArgumentKind.Value], _ => CellValue.Number(1)) { Status = FunctionStatus.NotImplemented });
        _wb.Functions = registry;
        Assert.Equal(Err(ErrorKind.Name), _wb.Evaluate("=LATER()"));
    }

    [Fact]
    public void Default_registry_has_sound_metadata()
    {
        // The full list is the compatibility report's job; this checks what every entry must satisfy.
        var all = FunctionRegistry.Default.All.ToList();
        Assert.Contains(all, f => f.Name == "SUM");
        Assert.All(all, f =>
        {
            Assert.Equal(f.Name.ToUpperInvariant(), f.Name);
            Assert.InRange(f.MinArguments, 0, f.MaxArguments);
            Assert.InRange(f.MaxArguments, 0, FunctionRegistry.MaxArguments);
            Assert.True(f.Status != FunctionStatus.KnownDeviation || !string.IsNullOrWhiteSpace(f.Deviation), $"{f.Name} needs a deviation text.");
        });
        Assert.True(all.Single(f => f.Name == "NOW").IsVolatile);
    }
}
