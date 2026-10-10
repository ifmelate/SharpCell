using System;
using System.Collections.Generic;

namespace SharpCell.Xlsx;

/// <summary>
/// A workbook was not saved because SharpCell could not calculate some of its formulas; see
/// <see cref="XlsxWriteOptions.KeepUncalculated"/>.
/// </summary>
public sealed class XlsxWriteException : InvalidOperationException
{
    internal XlsxWriteException(IReadOnlyList<CalculationDiagnostic> problems)
        : base($"{problems.Count} formula cell(s) could not be calculated, first {problems[0]}. "
            + $"Set {nameof(XlsxWriteOptions)}.{nameof(XlsxWriteOptions.KeepUncalculated)} to keep Excel's results for them.")
    {
        Problems = problems;
    }

    /// <summary>The formula cells that could not be calculated, with the reason for each.</summary>
    public IReadOnlyList<CalculationDiagnostic> Problems { get; }
}
