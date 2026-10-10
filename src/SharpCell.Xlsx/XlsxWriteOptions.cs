namespace SharpCell.Xlsx;

/// <summary>Options for <c>XlsxWriter</c>.</summary>
public sealed class XlsxWriteOptions
{
    /// <summary>The defaults: a save is refused when some formulas cannot be calculated.</summary>
    public XlsxWriteOptions()
    {
    }

    /// <summary>
    /// What to do with formulas SharpCell cannot calculate, such as calls of functions it does not
    /// know or links to other workbooks. False (the default) refuses to save with an
    /// <see cref="XlsxWriteException"/>. True keeps the result Excel saved for them and marks the file
    /// so that Excel recalculates every formula when it opens it; until then, cells that read such a
    /// formula show what SharpCell calculated for them.
    /// </summary>
    public bool KeepUncalculated { get; init; }

    /// <summary>Writes every value again, even one the file already has; tests the encoding on whole files.</summary>
    internal bool RewriteAllValues { get; init; }
}
