using System.Globalization;
using System.IO;
// snippet: sample-usings
using SharpCell;
using SharpCell.Xlsx;
// end-snippet

namespace SharpCell.Sample;

/// <summary>
/// Reads a small budget, recalculates it, changes an input, recalculates again and evaluates
/// formulas that live in no cell. Each step is a snippet in docs/getting-started.md.
/// </summary>
public static class BudgetDemo
{
    public static void Run(string path, TextWriter output)
    {
        // snippet: sample-load
        // Read the workbook: sheets, values, formulas and any results Excel saved with them.
        Workbook workbook = XlsxReader.Load(path);
        Worksheet budget = workbook["Budget"];

        // Calculate every formula now instead of relying on saved results.
        workbook.Recalculate();
        // end-snippet

        // snippet: sample-read
        CellValue total = budget["B6"].Value;   // =SUM(B2:B5)
        output.WriteLine($"Total per month: {Format(total)}");
        output.WriteLine($"Left after expenses: {Format(budget["F2"].Value)}");
        // end-snippet

        // snippet: sample-change
        // Change an input. Dependent formulas are updated by the next Recalculate, not before.
        budget["B2"].Value = 1350;
        workbook.Recalculate();
        output.WriteLine("After rent change to 1350:");
        output.WriteLine($"Total per month: {Format(budget["B6"].Value)}");
        output.WriteLine($"Left after expenses: {Format(budget["F2"].Value)}");
        // end-snippet

        // snippet: sample-evaluate
        // Evaluate formulas that live in no cell. They can read the workbook's sheets.
        CellValue food = workbook.Evaluate("=XLOOKUP(\"Food\", Budget!A2:A5, Budget!B2:B5)");
        CellValue average = workbook.Evaluate("=LET(items, Budget!B2:B5, SUM(items) / COUNT(items))");
        CellValue sequence = workbook.Evaluate("=SEQUENCE(3)");   // a dynamic array: {1;2;3}
        output.WriteLine($"Food per month: {Format(food)}");
        output.WriteLine($"Average expense: {Format(average)}");
        output.WriteLine($"Sequence: {sequence}");
        // end-snippet
    }

    // snippet: sample-format
    // A CellValue can be a number, text, logical, error, array or empty; check Kind before As*.
    private static string Format(CellValue value) => value.Kind switch
    {
        CellValueKind.Number => value.AsNumber().ToString(CultureInfo.InvariantCulture),
        CellValueKind.Error => value.AsError().ToText(),   // "#DIV/0!", "#N/A", ...
        _ => value.ToString(),
    };
    // end-snippet
}
