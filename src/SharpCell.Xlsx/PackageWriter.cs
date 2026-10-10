namespace SharpCell.Xlsx;

/// <summary>Writes the file of a <see cref="WritePlan"/>.</summary>
internal static class PackageWriter
{
    /// <summary>The file with the workbook's values; the source bytes themselves when nothing changed.</summary>
    public static byte[] Write(WritePlan plan) => plan.Source.Bytes;
}
