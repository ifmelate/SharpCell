using System.Collections.Generic;
using SharpCell.Evaluation;
using SharpCell.Parsing;

namespace SharpCell;

/// <summary>How a formula read from a file is calculated.</summary>
internal enum LoadedFormulaKind
{
    /// <summary>Saved without the dynamic array flag: the result is reduced by implicit intersection.</summary>
    Legacy,

    /// <summary>A dynamic array formula: an array result spills.</summary>
    Dynamic,

    /// <summary>An array formula (Ctrl+Shift+Enter): the result fills a fixed area.</summary>
    Array,
}

/// <summary>
/// Fills a new, empty sheet from a file in one pass, without the per-cell dependency bookkeeping
/// of <see cref="Cell"/>. Formulas keep the values the file cached until the next recalculation,
/// which calculates all of them. Cells must not be set twice; call <see cref="Complete"/> at the end.
/// </summary>
internal sealed class SheetLoader(Worksheet sheet)
{
    private readonly List<(CellKey Anchor, Area Area)> _arrays = [];

    // Dynamic formulas whose cached spill area did not fit the budget as the file states it.
    private readonly List<(CellKey Anchor, Area Area)> _oversized = [];

    public void SetValue(int row, int column, CellValue value)
    {
        if (value.Kind == CellValueKind.Empty)
            return;
        sheet.Store.GetOrCreate(row, column).Value = value;
    }

    /// <param name="formula">The parsed formula; cells of a shared formula can pass the same tree.</param>
    /// <param name="area">
    /// For <see cref="LoadedFormulaKind.Dynamic"/>, the area the cached result spilled over; for
    /// <see cref="LoadedFormulaKind.Array"/>, the area of the array formula. Includes the cell itself.
    /// </param>
    /// <returns>
    /// False, with nothing set, for an array formula whose area does not fit the workbook's spill
    /// budget (<see cref="Workbook.MaxSpillCells"/>). A dynamic formula whose stated spill area does
    /// not fit keeps the part the file actually has cells in, if that fits.
    /// </returns>
    /// <param name="row">Row of the formula cell, from 1.</param>
    /// <param name="column">Column of the formula cell, from 1.</param>
    /// <param name="kind">Whether the formula is ordinary, a dynamic array or a legacy array formula.</param>
    /// <param name="cached">The result the file saved, kept until the next recalculation.</param>
    public bool SetFormula(int row, int column, FormulaNode formula, LoadedFormulaKind kind, Area? area, CellValue cached)
    {
        var calculation = sheet.Workbook.Calculation;
        if (kind == LoadedFormulaKind.Array && area is { } fixedArea && !calculation.TryReserveLoadedArea(fixedArea))
            return false;

        var origin = new CellAddress(row, column);
        var data = Add(row, column, "=" + FormulaPrinter.Print(formula, origin), formula, cached);
        data.IsLegacy = kind == LoadedFormulaKind.Legacy;
        if (kind == LoadedFormulaKind.Legacy || area is not { } covered)
            return true;

        var key = new CellKey(sheet, row, column);
        if (kind == LoadedFormulaKind.Array)
        {
            data.FixedArray = covered;
            data.SpillArea = covered;
            calculation.RegisterFixedArray(key, covered);
            _arrays.Add((key, covered));
        }
        else if (!covered.IsSingleCell && calculation.TryReserveLoadedArea(covered))
        {
            data.SpillArea = covered;
            _arrays.Add((key, covered));
        }
        else if (!covered.IsSingleCell)
        {
            _oversized.Add((key, covered));
        }

        return true;
    }

    /// <summary>A formula that could not be parsed: it keeps its text and evaluates to <c>#NAME?</c>.</summary>
    public void SetUnsupportedFormula(int row, int column, string text, string reason, CellValue cached)
    {
        var formula = text.StartsWith('=') ? text : "=" + text;
        Add(row, column, formula, new UnsupportedNode(formula[1..], reason), cached);
    }

    /// <summary>Cells inside an anchor's area become its spilled cells, so they do not block it.</summary>
    public void Complete()
    {
        var calculation = sheet.Workbook.Calculation;
        foreach (var (anchor, claimed) in _oversized)
        {
            var used = Area.Cell(anchor.Row, anchor.Column);
            foreach (var cell in sheet.Store.Enumerate(claimed.FirstRow, claimed.FirstColumn, claimed.LastRow, claimed.LastColumn))
            {
                if (cell.Data.Formula is null)
                    used = Area.Bounding(used, Area.Cell(cell.Row, cell.Column));
            }

            if (!used.IsSingleCell && calculation.TryReserveLoadedArea(used) && anchor.Data is { } data)
            {
                data.SpillArea = used;
                _arrays.Add((anchor, used));
            }
        }

        _oversized.Clear();
        foreach (var (anchor, area) in _arrays)
        {
            foreach (var cell in sheet.Store.Enumerate(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn))
            {
                if (cell.Data.Formula is null && cell.Data.SpillAnchor is null)
                    cell.Data.SpillAnchor = anchor;
            }
        }

        _arrays.Clear();
    }

    private CellData Add(int row, int column, string text, FormulaNode formula, CellValue cached)
    {
        var data = sheet.Store.GetOrCreate(row, column);
        data.FormulaText = text;
        data.Formula = formula;
        data.Value = cached.Kind is CellValueKind.Missing or CellValueKind.Array or CellValueKind.Lambda ? CellValue.Empty : cached;
        sheet.Workbook.Calculation.MarkLoaded(new CellKey(sheet, row, column), data);
        return data;
    }
}
