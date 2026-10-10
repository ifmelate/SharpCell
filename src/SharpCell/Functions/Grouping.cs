using System;
using System.Collections.Generic;
using System.Linq;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// GROUPBY and PIVOTBY, as Excel for the web calculates them (tests/corpus/excel-web/groupby.xlsx):
/// rows are grouped by their field values, each group's values are passed to the function (as a
/// one-column array; PERCENTOF also gets the values it is a part of), and groups are sorted
/// ascending with empty keys last. Totals say "Total", or "Grand Total" when there are subtotals;
/// the label cells a total leaves free, and PIVOTBY cells no row falls in, are empty text.
/// </summary>
internal static class Grouping
{
    // GROUPBY(row_fields, values, function, [field_headers], [total_depth], [sort_order], [filter_array], [field_relationship])
    public static Operand GroupBy(FunctionCall call)
    {
        if (!Read(call, 0, out var fields, out var error) || !Read(call, 1, out var values, out error))
            return error;
        if (fields.GetLength(0) != values.GetLength(0))
            return Value();
        if (!ReadFunctions(call, 2, out var functions))
            return Value();
        if (!ReadOption(call, 3, -1, 0, 3, out var fieldHeaders)
            || !ReadOption(call, 4, 1, -fields.GetLength(1), fields.GetLength(1), out var depth)
            || !ReadOption(call, 7, 0, 0, 1, out var relationship)
            || (relationship == 1 && Math.Abs(depth) > 1))
            return Value();

        var headers = Headers.Of(fieldHeaders, values);
        if (!Filter(call, 6, fields.GetLength(0), headers.InInput, out var rows, out error))
            return error;

        var input = new Input(fields, null, values, headers, rows) { Culture = call.Context.Culture };
        var outputColumns = values.GetLength(1) * functions.Count;
        if (!ReadSort(call, 5, fields.GetLength(1), outputColumns, out var sort))
            return Value();

        var tree = Tree.Build(input.Fields, rows, input.FirstRow);
        var aggregates = new Aggregator(call.Context, input, functions);
        var lines = new List<CellValue[]>();
        var k = fields.GetLength(1);
        var showHeaders = headers.Show || functions.Count > 1;
        if (showHeaders)
        {
            var header = new CellValue[k + outputColumns];
            for (var i = 0; i < k; i++)
                header[i] = headers.Show ? headers.Field(fields, i, "Row Field") : Text("");
            for (var v = 0; v < values.GetLength(1); v++)
            {
                for (var f = 0; f < functions.Count; f++)
                {
                    header[k + v * functions.Count + f] = functions.Count > 1
                        ? Text(functions[f] is BuiltInFunction builtIn ? builtIn.Name : "")
                        : headers.Show ? headers.Field(values, v, "Value") : Text("");
                }
            }

            lines.Add(header);
        }

        CellValue[] Line(Node node, int level, string? label)
        {
            var line = new CellValue[k + outputColumns];
            var path = node.Path();
            for (var i = 0; i < k; i++)
                line[i] = label is not null ? (i == 0 ? Text(label) : Text("")) : i < level ? path[i] : Text("");
            var all = input.Rows;
            for (var v = 0; v < values.GetLength(1); v++)
            {
                for (var f = 0; f < functions.Count; f++)
                    line[k + v * functions.Count + f] = aggregates.Of(f, v, node.Rows, all);
            }

            return line;
        }

        var totalLabel = Math.Abs(depth) >= 2 ? "Grand Total" : "Total";
        if (depth < 0)
            lines.Add(Line(tree, 0, totalLabel));
        if (relationship == 1)
        {
            var leaves = tree.Leaves().ToList();
            leaves.Sort((a, b) => CompareFlat(a, b, sort, k, aggregates, input));
            foreach (var leaf in leaves)
                lines.Add(Line(leaf, k, null));
        }
        else
        {
            void Emit(Node node, int level)
            {
                foreach (var child in Sorted(node.Children, level + 1, sort, k, aggregates, input))
                {
                    var subtotal = level + 1 < k && level + 1 <= Math.Abs(depth) - 1;
                    if (child.Children.Count == 0)
                    {
                        lines.Add(Line(child, k, null));
                        continue;
                    }

                    if (subtotal && depth < 0)
                        lines.Add(Line(child, level + 1, null));
                    Emit(child, level + 1);
                    if (subtotal && depth > 0)
                        lines.Add(Line(child, level + 1, null));
                }
            }

            Emit(tree, 0);
        }

        if (depth > 0)
            lines.Add(Line(tree, 0, totalLabel));
        return Grid(lines);
    }

    // PIVOTBY(row_fields, col_fields, values, function, [field_headers], [row_total_depth], [row_sort_order],
    //   [col_total_depth], [col_sort_order], [filter_array], [relative_to])
    public static Operand PivotBy(FunctionCall call)
    {
        if (!Read(call, 0, out var rowFields, out var error) || !Read(call, 1, out var columnFields, out error)
            || !Read(call, 2, out var values, out error))
            return error;
        var count = rowFields.GetLength(0);
        if (columnFields.GetLength(0) != count || values.GetLength(0) != count || values.GetLength(1) != 1)
            return Value();
        if (!ReadFunctions(call, 3, out var functions) || functions.Count != 1)
            return Value();
        int kr = rowFields.GetLength(1), kc = columnFields.GetLength(1);
        if (!ReadOption(call, 4, -1, 0, 3, out var fieldHeaders)
            || !ReadOption(call, 5, 1, -kr, kr, out var rowDepth)
            || !ReadOption(call, 7, 1, -kc, kc, out var columnDepth)
            || !ReadOption(call, 10, 0, 0, 4, out var relativeTo))
            return Value();

        var headers = Headers.Of(fieldHeaders, values);
        if (!Filter(call, 9, count, headers.InInput, out var rows, out error))
            return error;
        if (!ReadSort(call, 6, kr, 1, out var rowSort) || !ReadSort(call, 8, kc, 1, out var columnSort))
            return Value();

        var input = new Input(rowFields, columnFields, values, headers, rows) { Culture = call.Context.Culture };
        var aggregates = new Aggregator(call.Context, input, functions);
        var rowTree = Tree.Build(rowFields, rows, input.FirstRow);
        var columnTree = Tree.Build(columnFields, rows, input.FirstRow);

        var rowLines = Layout(rowTree, kr, rowDepth, rowSort, aggregates, input);
        var columnLines = Layout(columnTree, kc, columnDepth, columnSort, aggregates, input);

        var lines = new List<CellValue[]>();
        var width = kr + columnLines.Count;
        if (headers.Show)
        {
            for (var i = 0; i < kc; i++)
            {
                var line = Blank(width);
                line[kr] = headers.Field(columnFields, i, "Column Field");
                lines.Add(line);
            }
        }

        for (var level = 0; level < kc; level++)
        {
            var line = Blank(width);
            for (var c = 0; c < columnLines.Count; c++)
                line[kr + c] = columnLines[c].Label(level, kc);
            lines.Add(line);
        }

        if (headers.Show)
        {
            var line = Blank(width);
            for (var i = 0; i < kr; i++)
                line[i] = headers.Field(rowFields, i, "Row Field");
            for (var c = 0; c < columnLines.Count; c++)
                line[kr + c] = headers.Field(values, 0, "Value");
            lines.Add(line);
        }

        foreach (var row in rowLines)
        {
            var line = Blank(width);
            for (var i = 0; i < kr; i++)
                line[i] = row.Label(i, kr);
            for (var c = 0; c < columnLines.Count; c++)
            {
                var column = columnLines[c];
                var subset = row.Node.Rows.Intersect(column.Node.Rows).ToList();
                if (subset.Count == 0)
                {
                    line[kr + c] = Text("");
                    continue;
                }

                var all = relativeTo switch
                {
                    0 => column.Node.Rows,
                    1 => row.Node.Rows,
                    2 => input.Rows,
                    3 => row.Node.Rows.Intersect((column.Node.Parent ?? column.Node).Rows).ToList(),
                    _ => (row.Node.Parent ?? row.Node).Rows.Intersect(column.Node.Rows).ToList(),
                };
                line[kr + c] = aggregates.Of(0, 0, subset, all);
            }

            lines.Add(line);
        }

        return Grid(lines);
    }

    // The rows (or columns) of a pivot: groups in order with their subtotals and the total.
    private static List<Entry> Layout(Node tree, int fields, int depth, List<(int Column, bool Descending)> sort, Aggregator aggregates, Input input)
    {
        var entries = new List<Entry>();
        var totalLabel = Math.Abs(depth) >= 2 ? "Grand Total" : "Total";
        if (depth < 0)
            entries.Add(new Entry(tree, 0, totalLabel));

        void Emit(Node node, int level)
        {
            foreach (var child in Sorted(node.Children, level + 1, sort, fields, aggregates, input))
            {
                if (child.Children.Count == 0)
                {
                    entries.Add(new Entry(child, fields, null));
                    continue;
                }

                var subtotal = level + 1 <= Math.Abs(depth) - 1;
                if (subtotal && depth < 0)
                    entries.Add(new Entry(child, level + 1, null));
                Emit(child, level + 1);
                if (subtotal && depth > 0)
                    entries.Add(new Entry(child, level + 1, null));
            }
        }

        Emit(tree, 0);
        if (depth > 0)
            entries.Add(new Entry(tree, 0, totalLabel));
        return entries;
    }

    // A row (or column) of a pivot: a group, a subtotal (level less than the fields) or a total.
    private sealed record Entry(Node Node, int Level, string? Total)
    {
        public CellValue Label(int index, int fields)
        {
            if (Total is not null)
                return Text(index == 0 ? Total : "");
            var path = Node.Path();
            return index < Level ? path[index] : Text("");
        }
    }

    private static IEnumerable<Node> Sorted(List<Node> siblings, int level, List<(int Column, bool Descending)> sort, int fields,
        Aggregator aggregates, Input input)
    {
        var ordered = siblings.ToList();
        ordered.Sort((a, b) =>
        {
            foreach (var (column, descending) in sort)
            {
                int order;
                if (column <= fields)
                {
                    if (column != level)
                        continue;
                    order = Compare(a.Key, b.Key, descending, input);
                }
                else
                {
                    var index = column - fields - 1;
                    order = Compare(aggregates.Column(index, a.Rows, input.Rows), aggregates.Column(index, b.Rows, input.Rows), descending, input);
                }

                if (order != 0)
                    return order;
            }

            return Compare(a.Key, b.Key, false, input);
        });
        return ordered;
    }

    // In a table (field_relationship 1), rows sort as a flat table: by the sort columns, then by every field.
    private static int CompareFlat(Node a, Node b, List<(int Column, bool Descending)> sort, int fields, Aggregator aggregates, Input input)
    {
        var pathA = a.Path();
        var pathB = b.Path();
        foreach (var (column, descending) in sort)
        {
            var order = column <= fields
                ? Compare(pathA[column - 1], pathB[column - 1], descending, input)
                : Compare(aggregates.Column(column - fields - 1, a.Rows, input.Rows), aggregates.Column(column - fields - 1, b.Rows, input.Rows), descending, input);
            if (order != 0)
                return order;
        }

        for (var i = 0; i < fields; i++)
        {
            var order = Compare(pathA[i], pathB[i], false, input);
            if (order != 0)
                return order;
        }

        return 0;
    }

    private static int Compare(CellValue a, CellValue b, bool descending, Input input) =>
        ArrayShapeSortFunctions.CompareForSort(a, b, descending, input.Culture);

    private static bool Read(FunctionCall call, int index, out CellValue[,] array, out CellValue error) =>
        ArrayArguments.TryRead(call[index], call.Context, out array, out error);

    // A function, or a row of them (HSTACK(SUM, AVERAGE)).
    private static bool ReadFunctions(FunctionCall call, int index, out List<LambdaValue> functions)
    {
        functions = [];
        var value = call.Value(index);
        var elements = value.Kind == CellValueKind.Array ? value.AsArray() : new[,] { { value } };
        foreach (var element in elements)
        {
            if (element.Kind != CellValueKind.Lambda)
                return false;
            functions.Add(element.AsLambda());
        }

        return functions.Count > 0;
    }

    // An optional whole number within a range; missing (or empty) gives the default.
    private static bool ReadOption(FunctionCall call, int index, int absent, int min, int max, out int option)
    {
        option = absent;
        if (!call.Has(index) || call.Value(index).Kind == CellValueKind.Empty)
            return true;
        var number = call.Integer(index);
        if (number.IsError || number.AsNumber() < min || number.AsNumber() > max)
            return false;
        option = (int)number.AsNumber();
        return true;
    }

    // sort_order: column numbers over the fields then the value columns, negative for descending; a
    // list may only name fields.
    private static bool ReadSort(FunctionCall call, int index, int fields, int valueColumns, out List<(int Column, bool Descending)> sort)
    {
        sort = [];
        if (!call.Has(index) || call.Value(index).Kind == CellValueKind.Empty)
            return true;
        var numbers = new List<long>();
        if (!ArrayArguments.TryReadIntegers(call, index, numbers, out _))
            return false;
        foreach (var number in numbers)
        {
            var column = (int)Math.Min(Math.Abs(number), int.MaxValue);
            if (column == 0 || column > fields + valueColumns || (numbers.Count > 1 && column > fields))
                return false;
            sort.Add((column, number < 0));
        }

        return true;
    }

    // The rows a filter keeps, after the header row; none kept is #VALUE!.
    private static bool Filter(FunctionCall call, int index, int count, bool headerInInput, out List<int> rows, out CellValue error)
    {
        var first = headerInInput ? 1 : 0;
        rows = [.. Enumerable.Range(first, count - first)];
        error = default;
        if (!call.Has(index))
        {
            if (rows.Count > 0)
                return true;
            error = Value();
            return false;
        }

        if (!Read(call, index, out var filter, out error))
            return false;
        if (filter.GetLength(1) != 1 || (filter.GetLength(0) != count && filter.GetLength(0) != count - first))
        {
            error = Value();
            return false;
        }

        var offset = filter.GetLength(0) == count ? 0 : first;
        var kept = new List<int>();
        foreach (var row in rows)
        {
            var flag = Coercion.ToBoolean(filter[row - offset, 0]);
            if (flag.IsError)
            {
                error = flag;
                return false;
            }

            if (flag.AsBoolean())
                kept.Add(row);
        }

        rows = kept;
        if (rows.Count > 0)
            return true;
        error = Value();
        return false;
    }

    private static CellValue[] Blank(int width)
    {
        var line = new CellValue[width];
        Array.Fill(line, Text(""));
        return line;
    }

    private static Operand Grid(List<CellValue[]> lines)
    {
        var width = lines.Count == 0 ? 0 : lines[0].Length;
        if (ArrayArguments.Allocate(lines.Count, width) is not { } grid)
            return CellValue.Error(ErrorKind.Num);
        for (var r = 0; r < lines.Count; r++)
        {
            for (var c = 0; c < width; c++)
                grid[r, c] = lines[r][c];
        }

        return CellValue.Array(grid);
    }

    private static CellValue Text(string text) => CellValue.Text(text);

    private static CellValue Value() => CellValue.Error(ErrorKind.Value);

    /// <summary>The inputs after the header row, with the rows a filter kept.</summary>
    private sealed class Input(CellValue[,] fields, CellValue[,]? columns, CellValue[,] values, Headers headers, List<int> rows)
    {
        public CellValue[,] Fields { get; } = fields;

        public CellValue[,]? Columns { get; } = columns;

        public CellValue[,] Values { get; } = values;

        public List<int> Rows { get; } = rows;

        public int FirstRow { get; } = headers.InInput ? 1 : 0;

        public System.Globalization.CultureInfo Culture { get; set; } = System.Globalization.CultureInfo.InvariantCulture;
    }

    /// <summary>
    /// field_headers: missing guesses (headers when the first value is text over a number), 0 none,
    /// 1 headers in the input not shown, 2 none in the input but generated, 3 shown.
    /// </summary>
    private sealed record Headers(bool InInput, bool Show, bool Generated)
    {
        public static Headers Of(int option, CellValue[,] values) => option switch
        {
            0 => new Headers(false, false, false),
            1 => new Headers(true, false, false),
            2 => new Headers(false, true, true),
            3 => new Headers(true, true, false),
            _ => new Headers(Guess(values), false, false),
        };

        private static bool Guess(CellValue[,] values) =>
            values.GetLength(0) >= 2 && values[0, 0].Kind == CellValueKind.Text && values[1, 0].Kind == CellValueKind.Number;

        /// <summary>A column's header: its input header, or a generated name such as "Row Field 1".</summary>
        public CellValue Field(CellValue[,] input, int column, string generated) =>
            Generated ? CellValue.Text($"{generated} {column + 1}") : InInput ? input[0, column] : CellValue.Text("");
    }

    /// <summary>A group: the rows whose fields share a key at each level up to this one.</summary>
    private sealed class Node(CellValue key, Node? parent)
    {
        public CellValue Key { get; } = key;

        public Node? Parent { get; } = parent;

        public List<int> Rows { get; } = [];

        public List<Node> Children { get; } = [];

        public CellValue[] Path()
        {
            var path = new List<CellValue>();
            for (var node = this; node.Parent is not null; node = node.Parent)
                path.Add(node.Key);
            path.Reverse();
            return [.. path];
        }

        public IEnumerable<Node> Leaves()
        {
            if (Children.Count == 0)
            {
                yield return this;
                yield break;
            }

            foreach (var child in Children)
            {
                foreach (var leaf in child.Leaves())
                    yield return leaf;
            }
        }
    }

    private static class Tree
    {
        // Rows grouped level by level; a group keeps the first key it met (text compares ignoring case).
        public static Node Build(CellValue[,] fields, List<int> rows, int firstRow)
        {
            var root = new Node(CellValue.Empty, null);
            var levels = fields.GetLength(1);
            foreach (var row in rows)
            {
                root.Rows.Add(row);
                var node = root;
                for (var level = 0; level < levels; level++)
                {
                    var key = fields[row, level];
                    var child = node.Children.Find(c => ArrayShapeSortFunctions.CompareForSort(c.Key, key, false, System.Globalization.CultureInfo.InvariantCulture) == 0
                        && c.Key.Kind == key.Kind);
                    if (child is null)
                    {
                        child = new Node(key, node);
                        node.Children.Add(child);
                    }

                    child.Rows.Add(row);
                    node = child;
                }
            }

            return root;
        }
    }

    /// <summary>Calls the functions on groups of values, once per group, function and value column.</summary>
    private sealed class Aggregator(EvaluationContext context, Input input, List<LambdaValue> functions)
    {
        private readonly Dictionary<(int, int, string, string), CellValue> _cache = [];

        public int Functions => functions.Count;

        /// <summary>The value of output column index (value column, then function) for a group.</summary>
        public CellValue Column(int index, List<int> rows, List<int> all) =>
            Of(index % functions.Count, index / functions.Count, rows, all);

        public CellValue Of(int function, int valueColumn, List<int> rows, List<int> all)
        {
            var key = (function, valueColumn, string.Join(',', rows), string.Join(',', all));
            if (_cache.TryGetValue(key, out var cached))
                return cached;
            var lambda = functions[function];
            var subset = Column(valueColumn, rows);
            Operand[] arguments = lambda is BuiltInFunction { Name: "PERCENTOF" } ? [subset, Column(valueColumn, all)] : [subset];
            CellValue result;
            if (!Lambdas.Takes(lambda, arguments.Length, context))
            {
                result = CellValue.Error(ErrorKind.Value);
            }
            else
            {
                result = Evaluator.ToValue(Lambdas.Invoke(lambda, arguments, context), context);
                if (result.Kind == CellValueKind.Array)
                    result = result.AsArray() is { Length: 1 } single ? single[0, 0] : CellValue.Error(ErrorKind.Calc);
                if (result.Kind == CellValueKind.Lambda)
                    result = CellValue.Error(ErrorKind.Calc);
            }

            _cache[key] = result;
            return result;
        }

        private CellValue Column(int valueColumn, List<int> rows)
        {
            var array = new CellValue[rows.Count, 1];
            for (var i = 0; i < rows.Count; i++)
                array[i, 0] = input.Values[rows[i], valueColumn];
            return CellValue.Array(array);
        }
    }
}
