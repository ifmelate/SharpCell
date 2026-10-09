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
    public void SetFormula(int row, int column, FormulaNode formula, LoadedFormulaKind kind, Area? area, CellValue cached)
    {
        var origin = new CellAddress(row, column);
        var data = Add(row, column, "=" + FormulaPrinter.Print(formula, origin), formula, cached);
        data.IsLegacy = kind == LoadedFormulaKind.Legacy;
        if (kind == LoadedFormulaKind.Legacy || area is not { } covered)
            return;

        var key = new CellKey(sheet, row, column);
        if (kind == LoadedFormulaKind.Array)
        {
            data.FixedArray = covered;
            data.SpillArea = covered;
            _arrays.Add((key, covered));
        }
        else if (!covered.IsSingleCell)
        {
            data.SpillArea = covered;
            _arrays.Add((key, covered));
        }
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
