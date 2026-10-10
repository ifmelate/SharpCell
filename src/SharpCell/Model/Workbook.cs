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

    /// <summary>An empty workbook with no sheets, the invariant culture and the 1900 date system.</summary>
    public Workbook()
    {
        Calculation = new Calculation(this);
    }

    private readonly List<Table> _tables = [];
    private readonly Dictionary<string, Table> _tablesByName = new(StringComparer.Ordinal);

    // Per sheet, the tables' areas keyed by their top-left cell: tables do not overlap, so a cell is
    // in at most one.
    private readonly Dictionary<Worksheet, RangeIndex> _tableAreas = [];
    private readonly Dictionary<CellKey, Table> _tablesByCorner = [];

    /// <summary>The tables of all sheets, in the order they were added.</summary>
    public IReadOnlyList<Table> Tables => _tables;

    /// <summary>Finds a table by name, ignoring case.</summary>
    /// <returns>Whether the table exists.</returns>
    public bool TryGetTable(string name, out Table? table)
    {
        ArgumentNullException.ThrowIfNull(name);
        return _tablesByName.TryGetValue(name.ToUpperInvariant(), out table);
    }

    /// <summary>The table that contains a cell, or null.</summary>
    internal Table? TableAt(Worksheet sheet, int row, int column)
    {
        if (!_tableAreas.TryGetValue(sheet, out var areas))
            return null;
        Table? found = null;
        areas.Query(row, column, corner => found = _tablesByCorner[corner]);
        return found;
    }

    internal Table RegisterTable(Table table)
    {
        var upper = table.Name.ToUpperInvariant();
        if (_tablesByName.ContainsKey(upper) || Names.ContainsInAnyScope(upper))
            throw new ArgumentException($"The name '{table.Name}' is already used by a table or a defined name.", "name");
        if (!_tableAreas.TryGetValue(table.Worksheet, out var areas))
            _tableAreas[table.Worksheet] = areas = new RangeIndex();
        Table? overlapped = null;
        areas.QueryOverlap(table.Area, corner => overlapped = _tablesByCorner[corner]);
        if (overlapped is not null)
            throw new ArgumentException($"The range overlaps table '{overlapped.Name}'.", "range");

        var key = new CellKey(table.Worksheet, table.Area.FirstRow, table.Area.FirstColumn);
        areas.Add(table.Area, key);
        _tablesByCorner[key] = table;
        _tables.Add(table);
        _tablesByName[upper] = table;

        // Formulas that named the table were #REF!; formulas inside it may use [Column] without a table name.
        Calculation.InvalidateName(upper);
        Calculation.InvalidateArea(table.Worksheet, table.Area);
        return table;
    }

    /// <summary>The sheets in order.</summary>
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

    /// <summary>Whether serial dates count from 1900 or from 1904, as set in the file.</summary>
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
    /// <see cref="Evaluate(string)"/> last until the next call of it.
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

    /// <summary>
    /// What a file reader keeps about the file the workbook came from, for writing it back; null for
    /// a workbook built in code. The engine does not look inside.
    /// </summary>
    internal object? Source { get; set; }


    /// <summary>Gets a sheet by name, ignoring case.</summary>
    public Worksheet this[string name] =>
        TryGetSheet(name, out var sheet) ? sheet! : throw new KeyNotFoundException($"No sheet named '{name}'.");

    /// <summary>Finds a sheet by name, ignoring case.</summary>
    /// <returns>Whether the sheet exists.</returns>
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

    /// <summary>Adds a sheet at the end.</summary>
    /// <exception cref="ArgumentException">The name is not a valid sheet name, or a sheet with that name exists.</exception>
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
        if (_tablesByName.ContainsKey(name.ToUpperInvariant()))
            throw new ArgumentException($"A table named '{name}' exists; tables and defined names share one set of names.", nameof(name));
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
    /// <c>#REF!</c> when the workbook has no sheets. The formula is in no table and has no row of
    /// its own: a table reference without a table name is <c>#REF!</c> and <c>[#This Row]</c> is
    /// <c>#VALUE!</c>.
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
