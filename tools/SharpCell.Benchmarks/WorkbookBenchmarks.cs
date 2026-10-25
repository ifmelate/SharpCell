using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using BenchmarkDotNet.Jobs;
using SharpCell.Xlsx;

namespace SharpCell.Benchmarks;

[MemoryDiagnoser]
public class WorkbookBenchmarks
{
    private string _path = "";
    private Workbook _calculated = null!;
    private int _edit;

    [Params(1_000, 10_000, 100_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _path = Path.Combine(Path.GetTempPath(), $"sharpcell-bench-{Rows}.xlsx");
        _calculated = SalesWorkbook.Build(Rows);
        _calculated.Recalculate();
        XlsxWriter.Save(_calculated, _path);
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_path))
            File.Delete(_path);
    }

    [Benchmark(Description = "Load .xlsx")]
    public Workbook Load() => XlsxReader.Load(_path);

    [Benchmark(Description = "Load .xlsx and recalculate")]
    public RecalculationResult LoadAndRecalculate() => XlsxReader.Load(_path).Recalculate();

    [Benchmark(Description = "Change one input and recalculate")]
    public RecalculationResult ChangeOneCell()
    {
        _calculated["Data"]["D2"].Value = ++_edit;
        return _calculated.Recalculate();
    }

    [Benchmark(Description = "Clone a calculated workbook")]
    public Workbook Clone() => _calculated.Clone();

    [Benchmark(Description = "Save .xlsx")]
    public long Save()
    {
        using var stream = new MemoryStream();
        XlsxWriter.Save(_calculated, stream);
        return stream.Length;
    }
}

/// <summary>Full calculation of a workbook whose formulas have never been calculated.</summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.Monitoring, iterationCount: 10, warmupCount: 1)]
public class FirstCalculationBenchmarks
{
    private Workbook _workbook = null!;

    [Params(1_000, 10_000, 100_000)]
    public int Rows { get; set; }

    [IterationSetup]
    public void Setup() => _workbook = SalesWorkbook.Build(Rows);

    [Benchmark(Description = "Calculate every formula once")]
    public RecalculationResult FirstCalculation() => _workbook.Recalculate();
}
