namespace SharpCell;

/// <summary>The workbook's date system: which day serial number 0 or 1 stands for.</summary>
public enum DateSystem
{
    /// <summary>Serial 1 is 1900-01-01 and the non-existent 1900-02-29 is serial 60, as in Excel for Windows.</summary>
    Date1900,

    /// <summary>Serial 0 is 1904-01-01, as in old Excel for Mac.</summary>
    Date1904,
}
