using System;

namespace SharpCell.Parsing;

/// <summary>The rows of a table a structured reference covers; <see cref="All"/> is all three parts.</summary>
[Flags]
internal enum TableRows
{
    Headers = 1,
    Data = 2,
    Totals = 4,
    ThisRow = 8,
    All = Headers | Data | Totals,
}

/// <summary>
/// A parsed table reference: <c>Sales[[#Data],[Jan]:[Mar]]</c> is table "Sales", rows
/// <see cref="TableRows.Data"/>, columns "Jan" to "Mar". A null <see cref="Table"/> means the table
/// the formula is in; a null <see cref="FirstColumn"/> means every column. Column names are
/// unescaped.
/// </summary>
internal sealed record StructuredReference(string? Table, TableRows Rows, string? FirstColumn, string? LastColumn);
