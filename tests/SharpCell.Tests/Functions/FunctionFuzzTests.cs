using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using SharpCell;
using SharpCell.Functions;

namespace SharpCell.Tests.Functions;

/// <summary>
/// Every registered function, called with random arguments of every kind, must return a value:
/// no exception caught at the call boundary (a FunctionFailure diagnostic), no hang. Arguments
/// are reproducible by seed.
/// </summary>
public class FunctionFuzzTests
{
    // Set SHARPCELL_FUZZ_ITERATIONS for a longer local run.
    private static readonly int CallsPerFunction =
        int.TryParse(Environment.GetEnvironmentVariable("SHARPCELL_FUZZ_ITERATIONS"), out var n) ? Math.Max(1, n / 10) : 60;

    private static readonly TimeSpan Budget = TimeSpan.FromSeconds(5);

    private static readonly string[] Pool =
    [
        "0", "1", "-1", "2", "0.5", "-0.5", "3.7", "10", "100", "1E+308", "-1E+308", "1E-308", "2147483648", "9.99E+307", "45301", "2958466",
        "\"\"", "\"abc\"", "\"1\"", "\"12:00\"", "\"2024-01-10\"", "\"a*\"", "\">5\"", "\"<>\"", "\"~\"", "\"Ab(c\"", "\"1+2i\"", "\"m\"", "\"0.0%\"",
        "TRUE", "FALSE", "#N/A", "#DIV/0!", "#VALUE!",
        "{1,2;3,4}", "{1;2;3}", "{\"a\",TRUE,#N/A}", "{0}",
        "A1", "A1:A3", "A1:C3", "B1:B2", "A:A", "1:1", "Z9", "A1:A1048576",
        "", // an empty argument
    ];

    // A scalar parameter given a whole column or row makes the engine call the function once per
    // cell, a million times: slow by design, not a hang, so those references go to the other kinds.
    private static readonly string[] WholeLines = ["A:A", "1:1", "A1:A1048576"];

    private static readonly string[] ScalarPool = [.. Pool.Where(item => !WholeLines.Contains(item))];

    public static TheoryData<string> Functions()
    {
        var data = new TheoryData<string>();
        foreach (var name in FunctionRegistry.Default.All.Select(f => f.Name).Order(StringComparer.Ordinal))
            data.Add(name);
        return data;
    }

    [Theory]
    [MemberData(nameof(Functions))]
    public void Function_returns_a_value_for_any_arguments(string name)
    {
        FunctionRegistry.Default.TryGet(name, out var function);
        var random = new Random(StableSeed(name));
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = 1;
        sheet["A2"].Value = "text";
        sheet["A3"].Value = true;
        sheet["B1"].Value = -2.5;
        sheet["C3"].Formula = "=1/0";

        var maxArguments = Math.Min(function!.MaxArguments, function.MinArguments + 3);
        var failures = new List<string>();
        for (var i = 0; i < CallsPerFunction && failures.Count < 5; i++)
        {
            var count = random.Next(function.MinArguments, maxArguments + 1);
            var arguments = new string[count];
            for (var a = 0; a < count; a++)
            {
                var pool = function.KindAt(a) == ArgumentKind.Value ? ScalarPool : Pool;
                arguments[a] = pool[random.Next(pool.Length)];
            }

            var formula = $"={name}({string.Join(',', arguments)})";

            var watch = Stopwatch.StartNew();
            try
            {
                using var timeout = new CancellationTokenSource(Budget);
                workbook.Evaluate(formula, timeout.Token);
            }
            catch (OperationCanceledException)
            {
                failures.Add($"{formula}: did not finish within {Budget.TotalSeconds} s");
                continue;
            }
            catch (FormulaParseException)
            {
                continue;
            }

            if (workbook.Diagnostics.FirstOrDefault(d => d.Kind == DiagnosticKind.FunctionFailure) is { } failure)
                failures.Add($"{formula}: {failure.Message}");
            else if (watch.Elapsed > Budget)
                failures.Add($"{formula}: took {watch.Elapsed.TotalSeconds:0.0} s");
        }

        Assert.True(failures.Count == 0, string.Join("\n", failures));
    }

    private static int StableSeed(string text)
    {
        var hash = 2166136261u;
        foreach (var c in text)
            hash = (hash ^ c) * 16777619u;
        return (int)(hash & 0x7fffffff);
    }
}
