using SharpCell;

namespace SharpCell.Tests.Evaluation;

/// <summary>Regressions from the stage 3 review.</summary>
public class Stage3ReviewTests
{
    private readonly Workbook _wb = new();
    private readonly Worksheet _s;

    public Stage3ReviewTests()
    {
        _s = _wb.AddSheet("S");
    }

    private static CellValue N(double value) => CellValue.Number(value);

    private static CellValue Arr(CellValue[,] values) => CellValue.Array(values);

    private void RecalculateWithin(TimeSpan limit)
    {
        using var timeout = new CancellationTokenSource(limit);
        _wb.Recalculate(timeout.Token);
    }

    [Fact]
    public void Anchors_reading_each_others_spill_are_a_loop_not_a_hang()
    {
        _s["A1"].Formula = "={1;2}+0*B2";
        _s["B1"].Formula = "={1;2}+0*A2";
        RecalculateWithin(TimeSpan.FromSeconds(5));

        Assert.Equal(N(0), _s["A1"].Value);
        Assert.Equal(N(0), _s["B1"].Value);
        Assert.Contains(_wb.Diagnostics, d => d.Kind == DiagnosticKind.CircularReference);

        _s["C9"].Value = 1;
        RecalculateWithin(TimeSpan.FromSeconds(5));
    }

    [Fact]
    public void Anchor_reading_its_own_spill_is_always_a_loop()
    {
        _s["X1"].Value = 1;
        _s["A1"].Formula = "={1;2}*X1+0*A2";
        RecalculateWithin(TimeSpan.FromSeconds(5));
        Assert.Equal(N(0), _s["A1"].Value);
        Assert.Single(_wb.Diagnostics);

        _s["X1"].Value = 2;
        RecalculateWithin(TimeSpan.FromSeconds(5));
        Assert.Equal(N(0), _s["A1"].Value);
        Assert.Single(_wb.Diagnostics);
    }

    [Fact]
    public void Readers_of_a_spill_become_dirty_with_its_anchor()
    {
        _s["X1"].Value = 1;
        _s["A1"].Formula = "={1;2}*X1";
        _s["C1"].Formula = "=A2*10";
        _s["C2"].Formula = "=C1+1";
        _wb.Recalculate();
        Assert.Equal(N(21), _s["C2"].Value);

        _s["X1"].Value = 5;
        Assert.Equal(N(101), _wb.Evaluate("=C2"));
    }

    [Fact]
    public void Random_sheets_with_spills_always_finish_calculating()
    {
        string[] templates = ["={{1;2}}+0*{0}", "={{1,2}}*0+{0}", "={0}+1", "=SUM({0}:{1})", "=IF({0}>1,{{1;2;3}},{1})", "5", ""];
        for (var seed = 0; seed < 150; seed++)
        {
            var random = new Random(seed);
            var wb = new Workbook();
            var ws = wb.AddSheet("S");
            for (var step = 0; step < 3; step++)
            {
                for (var i = 0; i < 6; i++)
                {
                    var cell = ws[random.Next(1, 5), random.Next(1, 5)];
                    var template = templates[random.Next(templates.Length)];
                    string Ref() => new CellAddress(random.Next(1, 5), random.Next(1, 5)).ToString();
                    var text = string.Format(template, Ref(), Ref());
                    if (text.Length == 0)
                        cell.Value = CellValue.Empty;
                    else if (text.StartsWith('='))
                        cell.Formula = text;
                    else
                        cell.Value = double.Parse(text);
                }

                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
                try
                {
                    wb.Recalculate(timeout.Token);
                }
                catch (OperationCanceledException)
                {
                    var cells = string.Join("; ", ws.Store.Enumerate(1, 1, 10, 10).Select(c =>
                        $"{new CellAddress(c.Row, c.Column)}={c.Data.FormulaText ?? c.Data.Value.ToString()}{(c.Data.SpillAnchor is { } a ? "<" + a.Address : "")}"));
                    Assert.Fail($"seed {seed} step {step}: {cells}");
                }
            }
        }
    }

    public static TheoryData<string, CellValue> RepeatedDirtyReads => new()
    {
        { "=MAP({1,2},LAMBDA(v,IF(ISERROR(A1),B1+0,v+A1)))", Arr(new CellValue[,] { { 6, 7 } }) },
        { "=BYROW({1;2},LAMBDA(r,IFERROR(A1+0,B1+0)))", Arr(new CellValue[,] { { 5 }, { 5 } }) },
        { "=MAKEARRAY(1,2,LAMBDA(r,c,IFERROR(A1+0,B1+0)))", Arr(new CellValue[,] { { 5, 5 } }) },
        { "=REDUCE(0,{1,2},LAMBDA(a,v,IF(ISERROR(a),B1,a+A1)))", 10 },
        { "=SCAN(0,{1,2},LAMBDA(a,v,IF(ISERROR(a),B1,a+A1)))", Arr(new CellValue[,] { { 5, 10 } }) },
        { "=LET(c,A1+0,c+IFERROR(c,B1+0))", 10 },
        { "=LET(c,A1+0,f,LAMBDA(x,IFERROR(x,B1+0)),c+1+0*f(c))", 6 },
        { "=(A1+0)+IFERROR(A1+0,B1+0)", 10 },
        { "=A1+INDIRECT(IF(ISERROR(A1),\"B1\",\"Z1\"))", 6 },
    };

    [Theory]
    [MemberData(nameof(RepeatedDirtyReads))]
    public void Repeated_reads_of_a_dirty_cell_never_choose_a_branch(string formula, CellValue expected)
    {
        _s["Z1"].Value = 1;
        _s["B1"].Formula = formula;
        _s["A1"].Formula = "=5";
        _wb.Recalculate();
        Assert.Equal(expected, expected.Kind == CellValueKind.Array ? _wb.Evaluate("=B1#") : _s["B1"].Value);
        Assert.Empty(_wb.Diagnostics);
    }

    [Fact]
    public void Named_formulas_do_not_see_the_callers_LET_names()
    {
        _wb.DefineName("Rate", "=0.05");
        _wb.DefineName("Calc", "=LAMBDA(a,a*Rate)");
        _wb.DefineName("Total", "=Rate*10");
        _wb.DefineName("Calc2", "=LAMBDA(a,a+x)");
        Assert.Equal(N(0.5), _wb.Evaluate("=LET(Rate,2,Calc(10))"));
        Assert.Equal(N(0.5), _wb.Evaluate("=LET(Rate,2,Total)"));
        Assert.Equal(CellValue.Error(ErrorKind.Name), _wb.Evaluate("=LET(x,3,Calc2(1))"));
    }

    [Fact]
    public void Lambda_recursion_can_be_cancelled()
    {
        _wb.DefineName("Twice", "=LAMBDA(n,IF(n=0,0,Twice(n-1)+Twice(n-1)))");
        _s["A1"].Formula = "=Twice(40)";
        using var cancel = new CancellationTokenSource(TimeSpan.FromMilliseconds(200));
        var watch = System.Diagnostics.Stopwatch.StartNew();
        Assert.Throws<OperationCanceledException>(() => _wb.Recalculate(cancel.Token));
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(5), watch.Elapsed.ToString());

        using var cancelEvaluate = new CancellationTokenSource(TimeSpan.FromMilliseconds(200));
        Assert.Throws<OperationCanceledException>(() => _wb.Evaluate("=Twice(40)", cancelEvaluate.Token));
    }

    [Fact]
    public void Interrupted_large_spill_leaves_no_orphans()
    {
        _s["A1"].Formula = "=MAKEARRAY(1000,500,LAMBDA(r,c,r))";
        using (var cancel = new CancellationTokenSource(TimeSpan.FromMilliseconds(30)))
        {
            try
            {
                _wb.Recalculate(cancel.Token);
            }
            catch (OperationCanceledException)
            {
            }
        }

        _s["A1"].Formula = "={1;2;3}";
        _wb.Recalculate();
        Assert.Equal(N(3), _s["A3"].Value);
        Assert.Equal(3, _s.Store.Count);
    }
}
