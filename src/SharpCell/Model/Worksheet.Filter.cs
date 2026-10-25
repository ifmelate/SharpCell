using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using SharpCell.Evaluation;
using SharpCell.Functions;
using SharpCell.Parsing;

namespace SharpCell;

// The sheet's filter: which rows its criteria show, as Excel works them out.
public sealed partial class Worksheet
{
    /// <summary>The sheet's filter, Excel's Data &gt; Filter; null when it has none. Tables have filters of their own, which this is not.</summary>
    public AutoFilter? AutoFilter { get; private set; }

    /// <summary>
    /// Turns on Excel's filter for a range, its first row holding the buttons, and returns it to
    /// filter columns with. A filter the sheet had is removed first, showing the rows it hid.
    /// </summary>
    /// <param name="range">An A1 range on this sheet with at least two rows, such as <c>A1:D20</c>; whole columns
    /// such as <c>A:D</c> are filtered down to the last row with something in it.</param>
    /// <exception cref="ArgumentException">The text is not such a range, or the range overlaps a table.</exception>
    /// <exception cref="InvalidOperationException">The sheet was removed.</exception>
    public AutoFilter SetAutoFilter(string range)
    {
        ArgumentNullException.ThrowIfNull(range);
        if (!ReferenceSyntax.TryParseA1Area(range, new CellAddress(1, 1), out var parsed) || parsed.Kind is not (AreaKind.Range or AreaKind.Columns))
            throw new ArgumentException($"'{range}' is not a range such as A1:D20.", nameof(range));
        var area = Area.Resolve(parsed, new CellAddress(1, 1));
        // Whole columns are filtered down to their last row with something in it, as Excel does.
        if (parsed.Kind == AreaKind.Columns)
            area = area with { LastRow = Math.Max(area.FirstRow + 1, Store.Bounds?.LastRow ?? 0) };
        if (area.Rows < 2)
            throw new ArgumentException("A filtered range needs a row of buttons and at least one row below it.", nameof(range));
        ThrowIfCannotChange();
        foreach (var table in Workbook.Tables)
        {
            if (table.Worksheet == this && table.Area.TryIntersect(area, out _))
                throw new ArgumentException($"The range overlaps table '{table.Name}', which has a filter of its own.", nameof(range));
        }

        RemoveAutoFilter();
        AutoFilter = new AutoFilter(this, area);
        return AutoFilter;
    }

    /// <summary>
    /// Turns the sheet's filter off, as clicking Excel's Filter again does: the rows it hid show,
    /// and the sheet is no longer in <see cref="FilterMode"/>. Nothing happens without a filter.
    /// </summary>
    /// <exception cref="InvalidOperationException">The sheet was removed.</exception>
    public void RemoveAutoFilter()
    {
        ThrowIfCannotChange();
        if (AutoFilter is not { } filter)
            return;
        if (filter.Columns.Values.Any(column => column.HasCriteria))
        {
            for (var row = filter.Area.FirstRow + 1; row <= filter.Area.LastRow && _hiddenRows.Count > 0; row++)
                _hiddenRows.Remove(row);
        }

        AutoFilter = null;
        _filterMode = false;
    }

    /// <summary>Sets the filter as a file has it, for the reader.</summary>
    internal AutoFilter LoadAutoFilter(Area area) => AutoFilter = new AutoFilter(this, area);

    internal void ApplyFilter(AutoFilter filter)
    {
        var area = filter.Area;
        filter.Limits.Clear();
        var tests = new List<Func<int, bool>>();
        foreach (var (column, criteria) in filter.Columns)
        {
            if (criteria.HasCriteria)
                tests.Add(Test(filter, column, criteria, area.FirstRow + 1, area.LastRow));
        }

        for (var row = area.FirstRow + 1; row <= area.LastRow; row++)
        {
            if (tests.TrueForAll(test => test(row)))
                _hiddenRows.Remove(row);
            else
                _hiddenRows.Add(row);
        }

        _filterMode = tests.Count > 0;
    }

    // The value a cell has now, calculated first when it is out of date.
    internal CellValue CurrentValue(int row, int column)
    {
        var data = Store.Get(row, column);
        return data is { Formula: not null, IsDirty: true }
            ? Workbook.Calculation.EvaluateDetached(
                new ReferenceNode(null, AreaRef.Cell(new CellRef(AxisRef.Absolute(row), AxisRef.Absolute(column)))),
                this, new CellAddress(row, column), CancellationToken.None)
            : data?.Value ?? CellValue.Empty;
    }

    private Func<int, bool> Test(AutoFilter filter, int column, ColumnFilter criteria, int first, int last)
    {
        var culture = Workbook.Culture;
        var compare = culture.CompareInfo;
        var values = new CellValue[last - first + 1];
        for (var row = first; row <= last; row++)
            values[row - first] = CurrentValue(row, column);
        CellValue Value(int row) => values[row - first];
        string Text(int row) => new Cell(this, row, column).Text;
        static bool IsBlank(CellValue value) => value.Kind == CellValueKind.Empty || (value.Kind == CellValueKind.Text && value.AsText().Length == 0);
        var numbers = values.Where(v => v.Kind == CellValueKind.Number).Select(v => v.AsNumber()).Order().ToArray();

        switch (criteria.Type)
        {
            case ColumnFilterType.Values:
            {
                var shown = new HashSet<string>(criteria.Values, StringComparer.Create(culture, ignoreCase: true));
                return row =>
                {
                    var value = Value(row);
                    if (IsBlank(value))
                        return criteria.Blanks;
                    if (shown.Contains(Text(row)))
                        return true;
                    return criteria.Dates.Count > 0 && value.Kind == CellValueKind.Number && FormatAt(row, column) is { ShowsDate: true }
                        && DateSerial.TryToDateTime(value.AsNumber(), Workbook.DateSystem, out var date)
                        && criteria.Dates.Any(group => group.Contains(date));
                };
            }

            case ColumnFilterType.Custom:
            {
                var conditions = criteria.Conditions.Select(condition => Condition(condition)).ToList();
                return row => criteria.MatchAll ? conditions.TrueForAll(c => c(row)) : conditions.Exists(c => c(row));
            }

            case ColumnFilterType.Top or ColumnFilterType.Bottom:
            {
                if (numbers.Length == 0)
                    return _ => false;
                var count = criteria.Percent ? Math.Max(1, (int)Math.Floor(numbers.Length * criteria.Count / 100)) : (int)criteria.Count;
                count = Math.Min(count, numbers.Length);
                var top = criteria.Type == ColumnFilterType.Top;
                var limit = top ? numbers[^count] : numbers[count - 1];
                filter.Limits[column] = (limit, null);
                return row => Value(row) is { Kind: CellValueKind.Number } value && (top ? value.AsNumber() >= limit : value.AsNumber() <= limit);
            }

            default:
                return Dynamic(filter, column, criteria.Dynamic, numbers, Value);
        }

        Func<int, bool> Condition(FilterCondition condition)
        {
            var isNumber = double.TryParse(condition.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number);
            bool Equal(int row)
            {
                var value = Value(row);
                if (IsBlank(value))
                    return condition.Value.Length == 0;
                if (isNumber && value.Kind == CellValueKind.Number)
                    return value.AsNumber() == number;
                return Wildcard.Matches(condition.Value, Text(row), culture);
            }

            return condition.Operator switch
            {
                FilterOperator.Equal => Equal,
                FilterOperator.NotEqual => row => !Equal(row),
                _ => row =>
                {
                    var value = Value(row);
                    int order;
                    if (isNumber)
                    {
                        if (value.Kind != CellValueKind.Number)
                            return false;
                        order = value.AsNumber().CompareTo(number);
                    }
                    else
                    {
                        if (value.Kind != CellValueKind.Text || IsBlank(value))
                            return false;
                        order = compare.Compare(value.AsText(), condition.Value, CompareOptions.IgnoreCase);
                    }

                    return condition.Operator switch
                    {
                        FilterOperator.GreaterThan => order > 0,
                        FilterOperator.GreaterThanOrEqual => order >= 0,
                        FilterOperator.LessThan => order < 0,
                        _ => order <= 0,
                    };
                },
            };
        }
    }

    private Func<int, bool> Dynamic(AutoFilter filter, int column, DynamicFilter dynamic, double[] numbers, Func<int, CellValue> value)
    {
        bool Number(int row, out double number)
        {
            var cell = value(row);
            number = cell.Kind == CellValueKind.Number ? cell.AsNumber() : 0;
            return cell.Kind == CellValueKind.Number;
        }

        if (dynamic is DynamicFilter.AboveAverage or DynamicFilter.BelowAverage)
        {
            if (numbers.Length == 0)
                return _ => false;
            var mean = numbers.Average();
            filter.Limits[column] = (mean, null);
            return dynamic == DynamicFilter.AboveAverage
                ? row => Number(row, out var n) && n > mean
                : row => Number(row, out var n) && n < mean;
        }

        var system = Workbook.DateSystem;
        if (dynamic >= DynamicFilter.Quarter1)
        {
            // A quarter or month of any year.
            var (from, to) = dynamic <= DynamicFilter.Quarter4
                ? (3 * (dynamic - DynamicFilter.Quarter1) + 1, 3 * (dynamic - DynamicFilter.Quarter1) + 3)
                : (dynamic - DynamicFilter.January + 1, dynamic - DynamicFilter.January + 1);
            return row => Number(row, out var n) && DateSerial.TryToDate(Math.Floor(n), system, out _, out var month, out _) && month >= from && month <= to;
        }

        var today = Workbook.Clock.GetLocalNow().Date;
        var (start, end) = Period(dynamic, today);
        double first = DateSerial.FromDateTime(start, system), next = DateSerial.FromDateTime(end, system);
        filter.Limits[column] = (first, next);
        return row => Number(row, out var n) && n >= first && n < next;
    }

    // The first day of a period around today and the day after its last.
    private static (DateTime Start, DateTime End) Period(DynamicFilter dynamic, DateTime today)
    {
        var week = today.AddDays(-(int)today.DayOfWeek);
        var month = new DateTime(today.Year, today.Month, 1);
        var quarter = new DateTime(today.Year, 3 * ((today.Month - 1) / 3) + 1, 1);
        var year = new DateTime(today.Year, 1, 1);
        return dynamic switch
        {
            DynamicFilter.Yesterday => (today.AddDays(-1), today),
            DynamicFilter.Today => (today, today.AddDays(1)),
            DynamicFilter.Tomorrow => (today.AddDays(1), today.AddDays(2)),
            DynamicFilter.LastWeek => (week.AddDays(-7), week),
            DynamicFilter.ThisWeek => (week, week.AddDays(7)),
            DynamicFilter.NextWeek => (week.AddDays(7), week.AddDays(14)),
            DynamicFilter.LastMonth => (month.AddMonths(-1), month),
            DynamicFilter.ThisMonth => (month, month.AddMonths(1)),
            DynamicFilter.NextMonth => (month.AddMonths(1), month.AddMonths(2)),
            DynamicFilter.LastQuarter => (quarter.AddMonths(-3), quarter),
            DynamicFilter.ThisQuarter => (quarter, quarter.AddMonths(3)),
            DynamicFilter.NextQuarter => (quarter.AddMonths(3), quarter.AddMonths(6)),
            DynamicFilter.LastYear => (year.AddYears(-1), year),
            DynamicFilter.ThisYear => (year, year.AddYears(1)),
            DynamicFilter.NextYear => (year.AddYears(1), year.AddYears(2)),
            _ => (year, today.AddDays(1)),
        };
    }

    // The filter's range and criteria follow inserted and deleted rows and columns; a filter whose
    // rows are all deleted goes.
    private void MoveAutoFilter(SheetEdit edit)
    {
        if (AutoFilter is not { } filter)
            return;
        if (edit.Map(filter.Area) is not { } area || area.Rows < 2)
        {
            AutoFilter = null;
            _filterMode = false;
            return;
        }

        filter.Area = area;
        if (!edit.Rows && edit.Covers(area.FirstRow, area.LastRow))
            filter.MoveColumns(column => edit.Map(column, CellAddress.MaxColumn));
    }
}
