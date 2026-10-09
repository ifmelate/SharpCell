using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using SharpCell.Evaluation;
using SharpCell.Functions;
using SharpCell.Parsing;

namespace SharpCell;

/// <summary>
/// A workbook: sheets, defined names and calculation settings. Not thread-safe; separate
/// workbooks can be used from different threads.
/// </summary>
public sealed class Workbook
{
    private readonly List<Worksheet> _sheets = [];
    private CultureInfo _culture = CultureInfo.InvariantCulture;
    private DateSystem _dateSystem;

    public Workbook()
    {
        Calculation = new Calculation(this);
    }

    public IReadOnlyList<Worksheet> Sheets => _sheets;

    /// <summary>
    /// Culture for text-to-number conversion, number-to-text conversion and text ordering.
    /// The process culture is never used. Defaults to the invariant culture.
    /// </summary>
    public CultureInfo Culture
    {
        get => _culture;
        set
        {
            _culture = value ?? throw new ArgumentNullException(nameof(value));
            Calculation.InvalidateAll();
        }
    }

    public DateSystem DateSystem
    {
        get => _dateSystem;
        set
        {
            _dateSystem = value;
            Calculation.InvalidateAll();
        }
    }

    /// <summary>
    /// Current calculation problems: circular references and failing functions. An entry of a cell
    /// disappears when the cell is edited or calculated without the problem; entries from
    /// <see cref="Evaluate"/> last until the next call of it.
    /// </summary>
    public IReadOnlyList<CalculationDiagnostic> Diagnostics => Calculation.Diagnostics;

    internal NameTable Names { get; } = new();

    internal Calculation Calculation { get; }

    internal FunctionRegistry Functions { get; set; } = FunctionRegistry.Default;

    /// <summary>Clock for NOW and TODAY; replaced in tests.</summary>
    internal TimeProvider Clock { get; set; } = TimeProvider.System;

    /// <summary>
    /// Cells all spills and array formulas may cover together, anchors included; a result that would
    /// go beyond is #SPILL!. Bounds the memory a small file can claim (about 400 bytes per cell).
    /// </summary>
    internal long MaxSpillCells { get; set; } = 1L << 22;

    /// <summary>Source for RAND and friends; replaced in tests.</summary>
    internal Random Random { get; set; } = Random.Shared;


    /// <summary>Gets a sheet by name, ignoring case.</summary>
    public Worksheet this[string name] =>
        TryGetSheet(name, out var sheet) ? sheet! : throw new KeyNotFoundException($"No sheet named '{name}'.");

    public bool TryGetSheet(string name, out Worksheet? sheet)
    {
        foreach (var candidate in _sheets)
        {
            if (string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                sheet = candidate;
                return true;
            }
        }

        sheet = null;
        return false;
    }

    public Worksheet AddSheet(string name)
    {
        ValidateSheetName(name);
        if (TryGetSheet(name, out _))
            throw new ArgumentException($"A sheet named '{name}' already exists.", nameof(name));

        var sheet = new Worksheet(this, name);
        _sheets.Add(sheet);

        // Formulas that pointed at a missing sheet of this name were #REF! and may resolve now.
        Calculation.InvalidateAll();
        return sheet;
    }

    /// <summary>Defines or replaces a name. A name with a <paramref name="scope"/> is visible only on that sheet.</summary>
    public void DefineName(string name, string formula, Worksheet? scope = null)
    {
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(formula);
        if (!IsValidName(name))
            throw new ArgumentException($"'{name}' is not a valid name.", nameof(name));
        if (scope is not null && scope.Workbook != this)
            throw new ArgumentException("The scope sheet belongs to another workbook.", nameof(scope));

        var text = formula.StartsWith('=') ? formula : "=" + formula;
        var node = FormulaParser.Parse(text, new CellAddress(1, 1));
        var upper = name.ToUpperInvariant();
        Names.Set(new NameDefinition(upper, text, node, scope));
        Calculation.InvalidateName(upper);
    }

    /// <summary>Defines a name read from a file whose formula cannot be parsed: using it gives <c>#NAME?</c>.</summary>
    internal void DefineUnsupportedName(string name, string formula, string reason, Worksheet? scope)
    {
        var text = formula.StartsWith('=') ? formula : "=" + formula;
        var upper = name.ToUpperInvariant();
        Names.Set(new NameDefinition(upper, text, new UnsupportedNode(text[1..], reason), scope));
        Calculation.InvalidateName(upper);
    }

    /// <summary>
    /// Calculates every formula that is out of date: those whose inputs changed since the last
    /// calculation and those using volatile functions (NOW, RAND).
    /// </summary>
    /// <exception cref="OperationCanceledException">The token was cancelled; finished cells keep their new values.</exception>
    public void Recalculate(CancellationToken cancellationToken = default) => Calculation.Recalculate(cancellationToken);

    /// <summary>
    /// Evaluates a formula that belongs to no cell, as if it were in cell A1 of the first sheet.
    /// Out-of-date cells it reads are calculated first. References without a sheet are
    /// <c>#REF!</c> when the workbook has no sheets.
    /// </summary>
    public CellValue Evaluate(string formula) => Evaluate(formula, CancellationToken.None);

    /// <inheritdoc cref="Evaluate(string)"/>
    /// <exception cref="OperationCanceledException">The token was cancelled.</exception>
    public CellValue Evaluate(string formula, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(formula);
        var origin = new CellAddress(1, 1);
        var node = FormulaParser.Parse(formula, origin);
        return Calculation.EvaluateDetached(node, _sheets.Count > 0 ? _sheets[0] : null, origin, cancellationToken);
    }


    // A name must read as a name in both reference styles: "A1", "R1C1", "R", "TRUE" are not names.
    internal static bool IsValidName(string name)
    {
        if (name.Length is 0 or > 255)
            return false;

        return LexesAsName(name, ReferenceStyle.A1) && LexesAsName(name, ReferenceStyle.R1C1);
    }

    private static bool LexesAsName(string name, ReferenceStyle style)
    {
        try
        {
            var tokens = Lexer.Tokenize(name, new CellAddress(1, 1), style);
            return tokens.Count == 2 && tokens[0].Kind == TokenKind.Name && tokens[0].Sheet is null
                && tokens[0].Length == name.Length;
        }
        catch (FormulaParseException)
        {
            return false;
        }
    }

    private static void ValidateSheetName(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (name.Length is 0 or > 31)
            throw new ArgumentException("A sheet name must have 1 to 31 characters.", nameof(name));
        if (name.AsSpan().IndexOfAny(":\\/?*[]") >= 0)
            throw new ArgumentException("A sheet name cannot contain : \\ / ? * [ or ].", nameof(name));
        if (name[0] == '\'' || name[^1] == '\'')
            throw new ArgumentException("A sheet name cannot start or end with an apostrophe.", nameof(name));
    }
}
