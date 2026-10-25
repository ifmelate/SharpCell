using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Running;
using SharpCell.Xlsx;

namespace SharpCell.Benchmarks;

internal static class Program
{
    public static void Main(string[] args)
    {
        if (args.Length > 0 && args[0] == "--check")
        {
            Check(args.Length > 1 ? int.Parse(args[1]) : 1_000);
            return;
        }

        if (args.Length > 0 && args[0] == "--time")
        {
            Time(args.Length > 1 ? int.Parse(args[1]) : 10_000);
            return;
        }

        if (args.Length > 0 && args[0] == "--time-io")
        {
            TimeIo(args.Length > 1 ? int.Parse(args[1]) : 10_000);
            return;
        }

        BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
    }

    // A few seconds of Stopwatch timing, for a quick before-and-after while changing code.
    // BenchmarkDotNet gives the numbers to publish; this only tells whether a change helped.
    private static void Time(int rows)
    {
        const int Runs = 5;
        Workbook workbook = SalesWorkbook.Build(rows);
        long allocated = GC.GetTotalAllocatedBytes();
        var watch = Stopwatch.StartNew();
        workbook.Recalculate();
        Console.WriteLine($"first calculation: {watch.ElapsedMilliseconds} ms, {(GC.GetTotalAllocatedBytes() - allocated) >> 20} MB allocated");

        var times = new List<long>();
        allocated = GC.GetTotalAllocatedBytes();
        for (int i = 1; i <= Runs; i++)
        {
            workbook["Data"]["D2"].Value = i;
            watch.Restart();
            workbook.Recalculate();
            times.Add(watch.ElapsedMilliseconds);
        }

        times.Sort();
        Console.WriteLine($"change one input and recalculate: median {times[Runs / 2]} ms, {(GC.GetTotalAllocatedBytes() - allocated) / Runs >> 20} MB allocated per run");
        Console.WriteLine($"Summary!C1 = {workbook["Summary"]["C1"].Value}, Data!I{rows + 1} = {workbook["Data"][$"I{rows + 1}"].Value}");
    }

    // Stopwatch timing of loading, cloning and saving a calculated workbook, like --time. Many runs,
    // because the first ten or so after start-up still run code the JIT has not finished optimizing.
    private static void TimeIo(int rows)
    {
        const int Runs = 25;
        Workbook calculated = SalesWorkbook.Build(rows);
        calculated.Recalculate();
        string path = Path.Combine(Path.GetTempPath(), $"sharpcell-time-io-{rows}.xlsx");
        try
        {
            XlsxWriter.Save(calculated, path);
            Measure("load .xlsx", () => XlsxReader.Load(path));
            Measure("clone a calculated workbook", () => calculated.Clone());
            Measure("save .xlsx", () =>
            {
                using var stream = new MemoryStream();
                XlsxWriter.Save(calculated, stream);
                return stream.Length;
            });
        }
        finally
        {
            File.Delete(path);
        }

        static void Measure(string name, Func<object> action)
        {
            // Two warm-up runs, and a collection before each measured one so it does not pay for the last.
            action();
            action();
            var times = new List<long>();
            var allocations = new List<long>();
            var watch = new Stopwatch();
            for (int i = 0; i < Runs; i++)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                long allocated = GC.GetTotalAllocatedBytes(precise: true);
                watch.Restart();
                GC.KeepAlive(action());
                times.Add(watch.ElapsedMilliseconds);
                allocations.Add(GC.GetTotalAllocatedBytes(precise: true) - allocated);
            }

            times.Sort();
            allocations.Sort();
            Console.WriteLine($"{name}: median {times[Runs / 2]} ms, {allocations[Runs / 2] >> 20} MB allocated" +
                $" (runs {times[0]}-{times[Runs - 1]} ms, {allocations[0] >> 20}-{allocations[Runs - 1] >> 20} MB)");
        }
    }

    // Builds the benchmark workbook once and shows that it calculates without errors.
    private static void Check(int rows)
    {
        Workbook workbook = SalesWorkbook.Build(rows);
        RecalculationResult result = workbook.Recalculate();
        int errors = workbook.Sheets
            .SelectMany(sheet => sheet.Cells)
            .Count(cell => cell.Value.Kind == CellValueKind.Error);

        Console.WriteLine($"rows: {rows}, cells: {workbook.Sheets.Sum(s => s.Cells.Count())}");
        Console.WriteLine($"changed by first calculation: {result.ChangedCells.Count}");
        Console.WriteLine($"cells with errors: {errors}");
        Console.WriteLine($"diagnostics: {workbook.Diagnostics.Count}");
        Console.WriteLine($"Data!I{rows + 1} (running total): {workbook["Data"][$"I{rows + 1}"].Value}");
        Console.WriteLine($"Summary!C1 (SUMIFS): {workbook["Summary"]["C1"].Value}");
        Console.WriteLine($"Summary!F1 (XLOOKUP): {workbook["Summary"]["F1"].Value}");
        Console.WriteLine($"Summary!H1 (SORT UNIQUE): {workbook["Summary"]["H1"].Value}");
        Console.WriteLine($"Summary!J1 (LET): {workbook["Summary"]["J1"].Value}");
    }
}
