using System;
using System.Collections.Generic;
using System.Globalization;
using SharpCell.Evaluation;
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

    public IReadOnlyList<Worksheet> Sheets => _sheets;

    /// <summary>
    /// Culture for text-to-number conversion, number-to-text conversion and text ordering.
    /// The process culture is never used. Defaults to the invariant culture.
    /// </summary>
    public CultureInfo Culture
    {
        get => _culture;
        set => _culture = value ?? throw new ArgumentNullException(nameof(value));
    }

    public DateSystem DateSystem { get; set; }

    internal NameTable Names { get; } = new();

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
        Names.Set(new NameDefinition(name.ToUpperInvariant(), text, node, scope));
    }

    /// <summary>
    /// Evaluates a formula that belongs to no cell, as if it were in cell A1 of the first sheet.
    /// References without a sheet are <c>#REF!</c> when the workbook has no sheets.
    /// </summary>
    public CellValue Evaluate(string formula)
    {
        ArgumentNullException.ThrowIfNull(formula);
        var origin = new CellAddress(1, 1);
        var node = FormulaParser.Parse(formula, origin);
        var context = new EvaluationContext(this, _sheets.Count > 0 ? _sheets[0] : null, origin);
        return Evaluator.EvaluateFormula(node, context);
    }

    internal void OnCellChanged(Worksheet sheet, int row, int column)
    {
    }

    // A name must read as a name in both reference styles: "A1", "R1C1", "R", "TRUE" are not names.
    private static bool IsValidName(string name)
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
