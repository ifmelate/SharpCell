using System;
using System.Globalization;
using System.IO;
using SharpCell;
using SharpCell.Xlsx;

namespace SharpCell.AotSmoke;

/// <summary>
/// A handful of formulas through every part of the engine that reflection or trimming could break:
/// the function registry, dynamic arrays, LAMBDA, number formats, regular expressions, cultures
/// and the xlsx reader. Exit code 0 when every result is as expected.
/// </summary>
internal static class Program
{
    private static int _failures;

    public static int Main(string[] args)
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");
        sheet["A1"].Value = 2;
        sheet["A2"].Formula = "=A1*3";
        sheet["B1"].Formula = "=SEQUENCE(3)";
        workbook.Recalculate();

        Expect("formula", sheet["A2"].Value, "6");
        Expect("spill", sheet["B3"].Value, "3");
        Expect("SUM", workbook.Evaluate("=SUM(1,2,3)"), "6");
        Expect("XLOOKUP", workbook.Evaluate("=XLOOKUP(2,{1,2,3},{\"a\",\"b\",\"c\"})"), "\"b\"");
        Expect("LAMBDA", workbook.Evaluate("=MAP({1,2},LAMBDA(x,x*10))"), "{10,20}");
        Expect("TEXT", workbook.Evaluate("=TEXT(1234.5,\"#,##0.00\")"), "\"1,234.50\"");
        Expect("REGEX", workbook.Evaluate("=REGEXEXTRACT(\"order 42\",\"[0-9]+\")"), "\"42\"");
        Expect("date", workbook.Evaluate("=DATE(2024,1,10)"), "45301");
        Expect("error", workbook.Evaluate("=1/0"), "#DIV/0!");
        var german = new Workbook { Culture = CultureInfo.GetCultureInfo("de-DE") };
        Expect("culture", german.Evaluate("=VALUE(\"1,5\")"), "1.5");

        if (args.Length > 0)
        {
            using var stream = File.OpenRead(args[0]);
            var budget = XlsxReader.Load(stream);
            budget.Recalculate();
            Expect("xlsx", budget["Budget"]["B6"].Value, "1810");
        }

        Console.WriteLine(_failures == 0 ? "AOT smoke test passed." : $"AOT smoke test: {_failures} failure(s).");
        return _failures == 0 ? 0 : 1;
    }

    private static void Expect(string name, CellValue actual, string expected)
    {
        if (actual.ToString() == expected)
            return;
        _failures++;
        Console.WriteLine($"{name}: expected {expected}, got {actual}");
    }
}
