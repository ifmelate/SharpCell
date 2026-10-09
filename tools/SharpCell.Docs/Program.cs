using System;

namespace SharpCell.Docs;

/// <summary>
/// Builds the documentation.
/// <code>
/// dotnet run --project tools/SharpCell.Docs -f net10.0 -- generate      write snippets, API reference, llms.txt
/// dotnet run --project tools/SharpCell.Docs -f net10.0 -- check         write nothing; exit 1 if anything is out of date
/// dotnet run --project tools/SharpCell.Docs -f net10.0 -- site DIR      render the HTML site into DIR
/// </code>
/// </summary>
internal static class Program
{
    public static int Main(string[] args)
    {
        try
        {
            var layout = RepositoryLayout.Find();
            switch (args)
            {
                case ["generate"]:
                    var files = Generator.Run(layout);
                    Generator.Write(layout, files);
                    Console.WriteLine($"Generated {files.Count} files.");
                    return 0;
                case ["check"]:
                    var differences = Generator.Differences(layout, Generator.Run(layout));
                    foreach (var difference in differences)
                        Console.Error.WriteLine(difference);
                    if (differences.Count > 0)
                        Console.Error.WriteLine("The documentation is out of date; run the tool with 'generate'.");
                    return differences.Count == 0 ? 0 : 1;
                default:
                    Console.Error.WriteLine("Usage: SharpCell.Docs generate | check | site DIR");
                    return 2;
            }
        }
        catch (DocsException ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 1;
        }
    }
}
