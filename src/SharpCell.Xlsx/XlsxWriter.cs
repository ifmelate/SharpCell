using System;
using System.IO;
using System.Xml;

namespace SharpCell.Xlsx;

/// <summary>
/// Saves a workbook read by <see cref="XlsxReader"/> back into its file with new values: change
/// input cells, call <see cref="Workbook.Recalculate"/>, save. Everything SharpCell does not model
/// (styles, charts, pivot tables, comments, macros) is copied from the original file unchanged, and so
/// is every cell whose value did not change. Only values can change: new or changed formulas, new
/// sheets, names or tables cannot be saved.
/// </summary>
public static class XlsxWriter
{
    private static readonly XlsxWriteOptions Defaults = new();

    /// <summary>Saves the workbook to a file, replacing it if it exists (the file it was read from too).</summary>
    /// <param name="workbook">A workbook read by <see cref="XlsxReader"/> and recalculated.</param>
    /// <param name="path">Path of the .xlsx file to write. The file is replaced only once the new one is complete.</param>
    /// <exception cref="NotSupportedException">The workbook was built in code, or something other than cell values changed since it was read.</exception>
    /// <exception cref="InvalidOperationException">Some formulas are not calculated: call <see cref="Workbook.Recalculate"/> first.</exception>
    /// <exception cref="XlsxWriteException">SharpCell could not calculate some formulas; see <see cref="XlsxWriteOptions.KeepUncalculated"/>.</exception>
    public static void Save(Workbook workbook, string path) => Save(workbook, path, Defaults);

    /// <summary>Saves the workbook to a stream, which is left open.</summary>
    /// <param name="workbook">A workbook read by <see cref="XlsxReader"/> and recalculated.</param>
    /// <param name="stream">Where the .xlsx content goes; it need not be seekable. Nothing is written when the save is refused.</param>
    /// <exception cref="NotSupportedException">The workbook was built in code, or something other than cell values changed since it was read.</exception>
    /// <exception cref="InvalidOperationException">Some formulas are not calculated: call <see cref="Workbook.Recalculate"/> first.</exception>
    /// <exception cref="XlsxWriteException">SharpCell could not calculate some formulas; see <see cref="XlsxWriteOptions.KeepUncalculated"/>.</exception>
    public static void Save(Workbook workbook, Stream stream) => Save(workbook, stream, Defaults);

    /// <summary>Saves the workbook to a file, replacing it if it exists (the file it was read from too).</summary>
    /// <param name="workbook">A workbook read by <see cref="XlsxReader"/> and recalculated.</param>
    /// <param name="path">Path of the .xlsx file to write. The file is replaced only once the new one is complete.</param>
    /// <param name="options">What to do with formulas SharpCell cannot calculate.</param>
    /// <exception cref="NotSupportedException">The workbook was built in code, or something other than cell values changed since it was read.</exception>
    /// <exception cref="InvalidOperationException">Some formulas are not calculated: call <see cref="Workbook.Recalculate"/> first.</exception>
    /// <exception cref="XlsxWriteException">SharpCell could not calculate some formulas and <paramref name="options"/> does not keep them.</exception>
    public static void Save(Workbook workbook, string path, XlsxWriteOptions options)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(path);
        ArgumentNullException.ThrowIfNull(options);
        var bytes = Write(workbook, options);

        // Written next to the target and moved over it: a failure never leaves half a file, and the
        // file the workbook was read from can be the target (its bytes are already in memory).
        var full = Path.GetFullPath(path);
        var temporary = Path.Combine(Path.GetDirectoryName(full)!, $".{Path.GetFileName(full)}.{Guid.NewGuid():N}.tmp");
        try
        {
            File.WriteAllBytes(temporary, bytes);
            File.Move(temporary, full, overwrite: true);
        }
        finally
        {
            File.Delete(temporary);
        }
    }

    /// <summary>Saves the workbook to a stream, which is left open.</summary>
    /// <param name="workbook">A workbook read by <see cref="XlsxReader"/> and recalculated.</param>
    /// <param name="stream">Where the .xlsx content goes; it need not be seekable. Nothing is written when the save is refused.</param>
    /// <param name="options">What to do with formulas SharpCell cannot calculate.</param>
    /// <exception cref="NotSupportedException">The workbook was built in code, or something other than cell values changed since it was read.</exception>
    /// <exception cref="InvalidOperationException">Some formulas are not calculated: call <see cref="Workbook.Recalculate"/> first.</exception>
    /// <exception cref="XlsxWriteException">SharpCell could not calculate some formulas and <paramref name="options"/> does not keep them.</exception>
    public static void Save(Workbook workbook, Stream stream, XlsxWriteOptions options)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(options);
        stream.Write(Write(workbook, options));
    }

    // The whole file is built in memory first, so a refusal or a damaged source writes nothing.
    private static byte[] Write(Workbook workbook, XlsxWriteOptions options)
    {
        var plan = WritePlan.Create(workbook, options);
        try
        {
            return PackageWriter.Write(plan);
        }
        catch (XmlException ex)
        {
            throw new InvalidDataException($"The workbook's file contains malformed XML: {ex.Message}", ex);
        }
    }
}
