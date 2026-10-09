using System.Collections.Generic;

namespace SharpCell.Functions;

/// <summary>Function name to implementation. Each category registers itself from its own file.</summary>
internal sealed class FunctionRegistry
{
    // Excel caps calls at 255 arguments.
    public const int MaxArguments = 255;

    private readonly Dictionary<string, FunctionInfo> _functions = [];

    public static FunctionRegistry Default { get; } = CreateDefault();

    public IEnumerable<FunctionInfo> All => _functions.Values;

    public void Add(FunctionInfo function) => _functions.Add(function.Name, function);

    /// <param name="upperName">Upper-case name without <c>_xlfn.</c> prefixes, as the parser produces.</param>
    public bool TryGet(string upperName, out FunctionInfo? function) => _functions.TryGetValue(upperName, out function);

    private static FunctionRegistry CreateDefault()
    {
        var registry = new FunctionRegistry();
        MathFunctions.Register(registry);
        LogicalFunctions.Register(registry);
        InformationFunctions.Register(registry);
        DateTimeFunctions.Register(registry);
        LambdaFunctions.Register(registry);
        ReferenceFunctions.Register(registry);
        LookupFunctions.Register(registry);
        ConditionalFunctions.Register(registry);
        return registry;
    }
}
