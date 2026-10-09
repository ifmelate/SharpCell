using System;
using System.IO;

namespace SharpCell.Sample;

internal static class Program
{
    public static int Main(string[] args)
    {
        var path = args.Length > 0 ? args[0] : Path.Combine(AppContext.BaseDirectory, "budget.xlsx");
        BudgetDemo.Run(path, Console.Out);
        return 0;
    }
}
