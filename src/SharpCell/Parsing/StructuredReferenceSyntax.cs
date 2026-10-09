using System.Text;

namespace SharpCell.Parsing;

/// <summary>
/// Parses the text of a structured reference token. Forms: <c>T[]</c>, <c>T[Col]</c>,
/// <c>T[#Totals]</c>, <c>T[@Col]</c>, <c>T[@[Col A]]</c>, and lists of bracketed parts such as
/// <c>T[[#Headers],[#Data],[A]:[B]]</c>. An apostrophe escapes the next character of a name.
/// The lexer has already checked that the brackets balance.
/// </summary>
internal static class StructuredReferenceSyntax
{
    /// <param name="text">The token text, ending in the bracket that closes the reference.</param>
    /// <param name="position">Where the token starts in the formula, for error positions.</param>
    public static StructuredReference Parse(string text, int position)
    {
        var open = text.IndexOf('[');
        var table = open == 0 ? null : text[..open];
        return new Body(text, open + 1, text.Length - 1, position).Parse(table);
    }

    // Parts that can be combined, as in Excel.
    private static bool Combines(TableRows rows) => rows is TableRows.Headers or TableRows.Data or TableRows.Totals
        or TableRows.ThisRow or TableRows.All or (TableRows.Headers | TableRows.Data) or (TableRows.Data | TableRows.Totals);

    private sealed class Body(string text, int start, int end, int offset)
    {
        private int _i = start;

        public StructuredReference Parse(string? table)
        {
            if (_i == end)
                return new StructuredReference(table, TableRows.Data, null, null);

            if (text[_i] == '@')
            {
                _i++;
                if (_i == end)
                    return new StructuredReference(table, TableRows.ThisRow, null, null);
                if (text[_i] != '[')
                {
                    var name = ReadName(bracketed: false);
                    return new StructuredReference(table, TableRows.ThisRow, name, name);
                }

                var (rows, first, last) = ReadParts();
                if (rows != 0)
                    throw Error("'@' cannot be combined with #All, #Data, #Headers, #Totals or #This Row.");
                return new StructuredReference(table, TableRows.ThisRow, first, last);
            }

            if (text[_i] != '[')
            {
                var (special, name) = ReadPart(bracketed: false);
                return special is { } item
                    ? new StructuredReference(table, item, null, null)
                    : new StructuredReference(table, TableRows.Data, name, name);
            }

            var parts = ReadParts();
            return new StructuredReference(table, parts.Rows == 0 ? TableRows.Data : parts.Rows, parts.First, parts.Last);
        }

        // [part], [part], ... where at most one part is a column or a column range [A]:[B].
        private (TableRows Rows, string? First, string? Last) ReadParts()
        {
            TableRows rows = 0;
            string? first = null, last = null;
            while (true)
            {
                Expect('[');
                var (special, name) = ReadPart(bracketed: true);
                if (special is { } item)
                {
                    if ((rows & item) != 0 || !Combines(rows | item))
                        throw Error("These parts of a table cannot be combined.");
                    rows |= item;
                }
                else
                {
                    if (first is not null)
                        throw Error("A structured reference names its columns once.");
                    first = last = name;
                    SkipSpaces();
                    if (_i < end && text[_i] == ':')
                    {
                        _i++;
                        SkipSpaces();
                        Expect('[');
                        var (endSpecial, endName) = ReadPart(bracketed: true);
                        if (endSpecial is not null)
                            throw Error("A column range must end in a column.");
                        last = endName;
                    }
                }

                SkipSpaces();
                if (_i == end)
                    return (rows, first, last);
                if (text[_i] != ',')
                    throw Error("Expected ',' between the parts of a structured reference.");
                _i++;
                SkipSpaces();
            }
        }

        private string ReadName(bool bracketed)
        {
            var (special, name) = ReadPart(bracketed);
            if (special is not null)
                throw Error("Expected a column name.");
            return name;
        }

        // A column name or a special item, up to the closing ']' (bracketed) or the end of the body.
        // An item is special only when its '#' is not escaped: '#Shame is the column "#Shame".
        private (TableRows? Special, string Name) ReadPart(bool bracketed)
        {
            var from = _i;
            var sb = new StringBuilder();
            var escapedFirst = false;
            while (true)
            {
                if (_i == end)
                {
                    if (bracketed)
                        throw Error("Expected ']'.");
                    break;
                }

                var c = text[_i];
                if (c == '\'')
                {
                    if (_i + 1 >= end)
                        throw Error("Expected a character after the escape \"'\".");
                    escapedFirst |= _i == from;
                    sb.Append(text[_i + 1]);
                    _i += 2;
                    continue;
                }

                if (c == '[')
                    throw Error("A '[' in a column name must be escaped as \"'[\".");
                if (c == ']')
                {
                    if (!bracketed)
                        throw Error("A ']' in a column name must be escaped as \"']\".");
                    _i++;
                    break;
                }

                sb.Append(c);
                _i++;
            }

            var content = sb.ToString();
            if (!escapedFirst && content.StartsWith('#'))
                return (Special(content), content);
            if (content.Length == 0)
                throw Error("A column name cannot be empty.");
            return (null, content);
        }

        private TableRows Special(string item) => item.ToUpperInvariant() switch
        {
            "#ALL" => TableRows.All,
            "#DATA" => TableRows.Data,
            "#HEADERS" => TableRows.Headers,
            "#TOTALS" => TableRows.Totals,
            "#THIS ROW" => TableRows.ThisRow,
            _ => throw Error($"'{item}' is not #All, #Data, #Headers, #Totals or #This Row."),
        };

        private void Expect(char c)
        {
            if (_i >= end || text[_i] != c)
                throw Error($"Expected '{c}'.");
            _i++;
        }

        private void SkipSpaces()
        {
            while (_i < end && text[_i] == ' ')
                _i++;
        }

        private FormulaParseException Error(string message) => new(message, offset + _i);
    }
}
