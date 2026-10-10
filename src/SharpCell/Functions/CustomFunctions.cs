using System;
using System.Collections.Generic;
using System.Globalization;
using SharpCell.Evaluation;
using SharpCell.Functions;
using SharpCell.Parsing;

namespace SharpCell;

/// <summary>
/// The functions a workbook's formulas can call besides Excel's: add one with <see cref="Add"/>,
/// and formulas calling it by name (ignoring case) use it from the next calculation on.
/// </summary>
public sealed class FunctionCollection
{
    private readonly Workbook _workbook;
    private readonly Dictionary<string, (string Name, FunctionInfo Info)> _functions = new(StringComparer.Ordinal);

    internal FunctionCollection(Workbook workbook) => _workbook = workbook;

    /// <summary>The names of the added functions, as they were added.</summary>
    public IReadOnlyCollection<string> Names
    {
        get
        {
            var names = new List<string>(_functions.Count);
            foreach (var (name, _) in _functions.Values)
                names.Add(name);
            return names;
        }
    }

    /// <summary>Whether a function of that name was added, ignoring case.</summary>
    public bool Contains(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        return _functions.ContainsKey(name.ToUpperInvariant());
    }

    /// <summary>
    /// Adds a function. Formulas that call the name, and were <c>#NAME?</c> so far, are calculated
    /// again by the next <see cref="Workbook.Recalculate"/>. Calls written <c>_xll.NAME(...)</c> or
    /// <c>_xludf.NAME(...)</c>, as Excel stores calls of add-in functions in files, find it too.
    /// </summary>
    /// <param name="name">Letters, digits, <c>.</c> and <c>_</c>, starting with a letter or <c>_</c>; not the
    /// name of a function SharpCell implements. A name Excel knows but SharpCell does not implement,
    /// such as <c>WEBSERVICE</c>, is allowed.</param>
    /// <param name="body">Calculates the result from the arguments. It must not change the workbook,
    /// and it should read cells only through its arguments: cells it reads otherwise are not
    /// dependencies, so it is not calculated again when they change. An exception it throws makes
    /// the cell <c>#VALUE!</c> with a <see cref="DiagnosticKind.CustomFunctionFailure"/> diagnostic.</param>
    /// <param name="options">Argument counts, which parameters are scalar, and whether the function
    /// is volatile; null for any number of arguments passed as they are.</param>
    /// <exception cref="ArgumentException">The name is not a function name, is built in, or was added already;
    /// or the options are not consistent.</exception>
    public void Add(string name, Func<FunctionArguments, CellValue> body, FunctionOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(body);
        _workbook.ThrowIfInCustomFunction();
        options ??= new FunctionOptions();
        var upper = ValidateName(name);
        if (_functions.ContainsKey(upper))
            throw new ArgumentException($"A function named '{name}' was added already; remove it first to replace it.", nameof(name));
        var kinds = options.Validate();

        var info = new FunctionInfo(upper, options.MinArguments, options.MaxArguments, kinds, call => Call(name, body, call))
        {
            IsVolatile = options.IsVolatile,
            IsCustom = true,
        };
        _functions.Add(upper, (name, info));
        InvalidateCalls(upper);
    }

    /// <summary>
    /// Removes a function. Formulas calling it become <c>#NAME?</c> at the next
    /// <see cref="Workbook.Recalculate"/>.
    /// </summary>
    /// <returns>Whether a function of that name was there.</returns>
    public bool Remove(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        _workbook.ThrowIfInCustomFunction();
        var upper = name.ToUpperInvariant();
        if (!_functions.Remove(upper))
            return false;
        InvalidateCalls(upper);
        return true;
    }

    /// <summary>The function a call names, prefixes of add-in calls removed; null when there is none.</summary>
    internal FunctionInfo? Find(string upperName)
    {
        if (_functions.Count == 0)
            return null;
        if (_functions.TryGetValue(upperName, out var found))
            return found.Info;
        foreach (var prefix in AddInPrefixes)
        {
            if (upperName.StartsWith(prefix, StringComparison.Ordinal) && _functions.TryGetValue(upperName[prefix.Length..], out found))
                return found.Info;
        }

        return null;
    }

    /// <summary>Copies the functions of another workbook into this empty collection; the bodies are shared.</summary>
    internal void CopyFrom(FunctionCollection source)
    {
        foreach (var (key, entry) in source._functions)
            _functions.Add(key, entry);
    }

    private static readonly string[] AddInPrefixes = ["_XLL.", "_XLUDF."];

    // Calls of the name were recorded as name reads (an unknown function, or this one), so marking
    // the name's readers out of date reaches them, prefixed spellings included.
    private void InvalidateCalls(string upper)
    {
        _workbook.Calculation.InvalidateName(upper);
        foreach (var prefix in AddInPrefixes)
            _workbook.Calculation.InvalidateName(prefix + upper);
    }

    private string ValidateName(string name)
    {
        if (name.Length is 0 or > 255 || !(char.IsAsciiLetter(name[0]) || name[0] == '_'))
            throw new ArgumentException($"'{name}' is not a function name: it must start with a letter or '_'.", nameof(name));
        foreach (var c in name)
        {
            if (!(char.IsAsciiLetterOrDigit(c) || c is '.' or '_'))
                throw new ArgumentException($"'{name}' is not a function name: use letters, digits, '.' and '_'.", nameof(name));
        }

        var upper = name.ToUpperInvariant();
        if (upper.StartsWith("_XL", StringComparison.Ordinal) && upper.Contains('.', StringComparison.Ordinal))
            throw new ArgumentException($"'{name}' starts with a prefix Excel reserves for files.", nameof(name));

        // The parser must read "NAME(" as a call of exactly this name.
        try
        {
            if (FormulaParser.Parse("=" + name + "()", new CellAddress(1, 1)) is not FunctionNode { Name: var parsed } || parsed != upper)
                throw new ArgumentException($"'{name}' does not read as a function name in a formula.", nameof(name));
        }
        catch (FormulaParseException)
        {
            throw new ArgumentException($"'{name}' does not read as a function name in a formula.", nameof(name));
        }

        if (upper is "LET" or "LAMBDA"
            || (_workbook.Registry.TryGet(upper, out var builtIn) && builtIn!.Status != FunctionStatus.NotImplemented))
            throw new ArgumentException($"{upper} is a function SharpCell implements; it cannot be replaced.", nameof(name));
        return upper;
    }

    private static Operand Call(string name, Func<FunctionArguments, CellValue> body, FunctionCall call)
    {
        var context = call.Context;
        var arguments = new FunctionArguments(call);
        var workbook = context.Workbook;
        CellValue result;
        workbook.CustomFunctionDepth++;
        try
        {
            result = body(arguments);
        }
        catch (OperationCanceledException) when (context.CancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception ex)
        {
            context.Report(DiagnosticKind.CustomFunctionFailure, $"{name} failed: {ex.GetType().Name}: {ex.Message}");
            return CellValue.Error(ErrorKind.Value);
        }
        finally
        {
            workbook.CustomFunctionDepth--;
            arguments.Close();
        }

        return result.Kind is CellValueKind.Missing or CellValueKind.Lambda ? CellValue.Error(ErrorKind.Value) : result;
    }
}

/// <summary>How a function added through <see cref="Workbook.Functions"/> takes its arguments.</summary>
public sealed class FunctionOptions
{
    /// <summary>Options for any number of arguments, all passed as they are, and a function that is not volatile.</summary>
    public FunctionOptions()
    {
    }

    /// <summary>The fewest arguments a call may have; fewer make the call <c>#VALUE!</c>. Default 0.</summary>
    public int MinArguments { get; init; }

    /// <summary>The most arguments a call may have, up to 255 as in Excel; more make the call <c>#VALUE!</c>. Default 255.</summary>
    public int MaxArguments { get; init; } = FunctionRegistry.MaxArguments;

    /// <summary>
    /// The 0-based positions of scalar parameters. A scalar parameter receives one value: given a
    /// range or an array, the function is called once per element and the results spill, as with
    /// Excel's own functions such as <c>LEN(A1:A10)</c>. Other parameters receive the argument as
    /// it is: a range as an array of its values.
    /// </summary>
    public IReadOnlyList<int> ScalarParameters { get; init; } = [];

    /// <summary>
    /// Whether the function is calculated on every <see cref="Workbook.Recalculate"/>, like
    /// <c>NOW</c>, instead of only when its arguments change.
    /// </summary>
    public bool IsVolatile { get; init; }

    internal ArgumentKind[] Validate()
    {
        if (MinArguments < 0 || MaxArguments > FunctionRegistry.MaxArguments || MinArguments > MaxArguments)
            throw new ArgumentException($"Argument counts must satisfy 0 <= MinArguments <= MaxArguments <= {FunctionRegistry.MaxArguments}.", "options");
        ArgumentNullException.ThrowIfNull(ScalarParameters, "options.ScalarParameters");

        // The last kind repeats for later arguments, so the array ends with a non-scalar entry.
        var last = -1;
        foreach (var index in ScalarParameters)
        {
            if (index < 0 || index >= MaxArguments)
                throw new ArgumentException($"Scalar parameter {index} is outside 0 to {MaxArguments - 1}.", "options");
            last = Math.Max(last, index);
        }

        var kinds = new ArgumentKind[last + 2];
        Array.Fill(kinds, ArgumentKind.Any);
        foreach (var index in ScalarParameters)
            kinds[index] = ArgumentKind.Value;
        return kinds;
    }
}

/// <summary>
/// The arguments of one call of a function added through <see cref="Workbook.Functions"/>. Valid
/// only while the function runs.
/// </summary>
public sealed class FunctionArguments
{
    private FunctionCall? _call;

    internal FunctionArguments(FunctionCall call) => _call = call;

    private FunctionCall Call => _call ?? throw new InvalidOperationException("The arguments are valid only while the function runs.");

    /// <summary>How many arguments the call has, omitted ones included.</summary>
    public int Count => Call.Count;

    /// <summary>
    /// The argument's value: a single cell's value, a range as an <see cref="CellValueKind.Array"/>,
    /// an error as an error value, and <see cref="CellValue.Missing"/> for an argument left out, as
    /// the middle one of <c>F(1,,2)</c>. An array is a copy the function may change.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no such argument.</exception>
    public CellValue this[int index]
    {
        get
        {
            CheckIndex(index);
            if (Call.IsMissing(index))
                return CellValue.Missing;

            // An array constant is the formula's own, shared with clones of the workbook.
            var value = Call.Value(index);
            return value.Kind == CellValueKind.Array ? CellValue.Array((CellValue[,])value.AsArray().Clone()) : value;
        }
    }

    /// <summary>Whether the argument was left out, as the middle one of <c>F(1,,2)</c>.</summary>
    public bool IsMissing(int index)
    {
        CheckIndex(index);
        return Call.IsMissing(index);
    }

    /// <summary>
    /// The argument converted to a number as Excel converts it (text such as <c>"2"</c> included; an
    /// omitted argument is 0): a number, or the error to return, such as <c>#VALUE!</c> for <c>"abc"</c>.
    /// </summary>
    public CellValue Number(int index)
    {
        CheckIndex(index);
        return Call.Number(index);
    }

    /// <summary>The argument converted to text as Excel converts it: text, or the error to return.</summary>
    public CellValue Text(int index)
    {
        CheckIndex(index);
        return Call.Text(index);
    }

    /// <summary>The argument converted to a logical value as Excel converts it: TRUE or FALSE, or the error to return.</summary>
    public CellValue Boolean(int index)
    {
        CheckIndex(index);
        return Call.Boolean(index);
    }

    /// <summary>The cell whose formula makes the call; null when the formula was given to <see cref="Workbook.Evaluate(string)"/>.</summary>
    public Cell? Caller
    {
        get
        {
            var context = Call.Context;
            return context.IsDetached || context.Sheet is null ? null : new Cell(context.Sheet, context.Origin.Row, context.Origin.Column);
        }
    }

    /// <summary>The workbook's date system, for reading date arguments with <see cref="CellValue.AsDateTime"/>.</summary>
    public DateSystem DateSystem => Call.Context.DateSystem;

    /// <summary>The workbook's culture, as used for converting between text and numbers.</summary>
    public CultureInfo Culture => Call.Context.Culture;

    /// <summary>
    /// The token passed to <see cref="Workbook.Recalculate"/> or <see cref="Workbook.Evaluate(string, System.Threading.CancellationToken)"/>,
    /// for a slow function to stop early: an <see cref="OperationCanceledException"/> it throws once
    /// the token is cancelled cancels the calculation.
    /// </summary>
    public System.Threading.CancellationToken CancellationToken => Call.Context.CancellationToken;

    internal void Close() => _call = null;

    private void CheckIndex(int index)
    {
        if ((uint)index >= (uint)Call.Count)
            throw new ArgumentOutOfRangeException(nameof(index), index, $"The call has {Call.Count} arguments.");
    }
}
