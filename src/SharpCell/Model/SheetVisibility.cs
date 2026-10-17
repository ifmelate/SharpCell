namespace SharpCell;

/// <summary>Whether a sheet's tab shows in Excel.</summary>
public enum SheetVisibility
{
    /// <summary>The tab shows.</summary>
    Visible,

    /// <summary>Hidden; Excel's Unhide shows it again.</summary>
    Hidden,

    /// <summary>Hidden so that Excel's Unhide does not list it; only code (VBA, SharpCell) shows it again.</summary>
    VeryHidden,
}
