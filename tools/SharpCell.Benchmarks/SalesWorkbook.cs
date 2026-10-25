using System;

namespace SharpCell.Benchmarks;

/// <summary>
/// Builds a workbook shaped like a typical sales model: a data sheet with formulas on every row,
/// running totals, and a summary sheet that aggregates and looks up over the whole data range.
/// The content is a function of <c>rows</c> alone, so every run measures the same workbook.
/// </summary>
internal static class SalesWorkbook
{
    private const int Regions = 5;
    private const int Products = 20;
    private const int RunLength = 100;

    public static Workbook Build(int rows)
    {
        var workbook = new Workbook();
        Worksheet data = workbook.AddSheet("Data");
        Worksheet summary = workbook.AddSheet("Summary");
        var random = new Random(12345);

        string[] header = ["Id", "Region", "Product", "Units", "Price", "Date", "Revenue", "Size", "Running"];
        for (int c = 0; c < header.Length; c++)
            data[1, c + 1].Value = header[c];

        for (int i = 0; i < rows; i++)
        {
            int r = i + 2;
            data[$"A{r}"].Value = i + 1;
            data[$"B{r}"].Value = "Region " + (i % Regions + 1);
            data[$"C{r}"].Value = "Product " + (random.Next(Products) + 1);
            data[$"D{r}"].Value = random.Next(1, 200);
            data[$"E{r}"].Value = Math.Round(random.NextDouble() * 90 + 10, 2);
            data[$"F{r}"].Value = 45292 + random.Next(365);
            data[$"G{r}"].Formula = $"=D{r}*E{r}";
            data[$"H{r}"].Formula = $"=IF(G{r}>5000,\"big\",\"small\")";
            data[$"I{r}"].Formula = r == 2
                ? $"=G{r}"
                : $"=IF(MOD(A{r},{RunLength})=1,G{r},I{r - 1}+G{r})";
        }

        int last = rows + 1;
        for (int region = 1; region <= Regions; region++)
        {
            for (int product = 1; product <= Products; product++)
            {
                int r = (region - 1) * Products + product;
                summary[$"A{r}"].Value = "Region " + region;
                summary[$"B{r}"].Value = "Product " + product;
                summary[$"C{r}"].Formula =
                    $"=SUMIFS(Data!G2:G{last},Data!B2:B{last},A{r},Data!C2:C{last},B{r})";
            }
        }

        for (int i = 0; i < 200; i++)
        {
            int r = i + 1;
            summary[$"E{r}"].Value = (i * 7) % Math.Max(rows, 1) + 1;
            summary[$"F{r}"].Formula = $"=XLOOKUP(E{r},Data!A2:A{last},Data!G2:G{last},0)";
        }

        summary["H1"].Formula = $"=SORT(UNIQUE(Data!C2:C{last}))";
        summary["J1"].Formula = $"=LET(total, SUM(Data!G2:G{last}), total / COUNTA(Data!A2:A{last}))";
        return workbook;
    }
}
