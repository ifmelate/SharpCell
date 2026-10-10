using System;
using System.IO;
using System.Xml;

namespace SharpCell.Xlsx;

/// <summary>
/// Saves a workbook as an .xlsx file: values, formulas, styles, number formats, sizes, merged cells,
/// panes, sheets, defined names and tables. A workbook read by <see cref="XlsxReader"/> is saved into
/// the file it was read from: everything SharpCell does not model (charts, pivot tables, comments,
/// conditional formats, macros) is copied unchanged, and so is every cell that did not change. A
/// workbook built in code becomes a new file. Call <see cref="Workbook.Recalculate"/> before saving.
/// </summary>
public static class XlsxWriter
{
    private static readonly XlsxWriteOptions Defaults = new();

    /// <summary>Saves the workbook to a file, replacing it if it exists (the file it was read from too).</summary>
    /// <param name="workbook">A recalculated workbook, read by <see cref="XlsxReader"/> or built in code.</param>
    /// <param name="path">Path of the .xlsx file to write (.xlsm for a macro-enabled workbook read from one). The file is replaced only once the new one is complete.</param>
    /// <exception cref="NotSupportedException">A cell of an Excel data table changed, or a cell holds <c>#SPILL!</c> or <c>#CALC!</c> as a value.</exception>
    /// <exception cref="InvalidOperationException">Some formulas are not calculated: call <see cref="Workbook.Recalculate"/> first; or the workbook has no sheets.</exception>
    /// <exception cref="XlsxWriteException">SharpCell could not calculate some formulas; see <see cref="XlsxWriteOptions.KeepUncalculated"/>.</exception>
    public static void Save(Workbook workbook, string path) => Save(workbook, path, Defaults);

    /// <summary>Saves the workbook to a stream, which is left open.</summary>
    /// <param name="workbook">A recalculated workbook, read by <see cref="XlsxReader"/> or built in code.</param>
    /// <param name="stream">Where the .xlsx content goes; it need not be seekable. Nothing is written when the save is refused.</param>
    /// <exception cref="NotSupportedException">A cell of an Excel data table changed, or a cell holds <c>#SPILL!</c> or <c>#CALC!</c> as a value.</exception>
    /// <exception cref="InvalidOperationException">Some formulas are not calculated: call <see cref="Workbook.Recalculate"/> first; or the workbook has no sheets.</exception>
    /// <exception cref="XlsxWriteException">SharpCell could not calculate some formulas; see <see cref="XlsxWriteOptions.KeepUncalculated"/>.</exception>
    public static void Save(Workbook workbook, Stream stream) => Save(workbook, stream, Defaults);

    /// <summary>Saves the workbook to a file, replacing it if it exists (the file it was read from too).</summary>
    /// <param name="workbook">A recalculated workbook, read by <see cref="XlsxReader"/> or built in code.</param>
    /// <param name="path">Path of the .xlsx file to write (.xlsm for a macro-enabled workbook read from one). The file is replaced only once the new one is complete.</param>
    /// <param name="options">What to do with formulas SharpCell cannot calculate.</param>
    /// <exception cref="NotSupportedException">A cell of an Excel data table changed, or a cell holds <c>#SPILL!</c> or <c>#CALC!</c> as a value.</exception>
    /// <exception cref="InvalidOperationException">Some formulas are not calculated: call <see cref="Workbook.Recalculate"/> first; or the workbook has no sheets.</exception>
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
    /// <param name="workbook">A recalculated workbook, read by <see cref="XlsxReader"/> or built in code.</param>
    /// <param name="stream">Where the .xlsx content goes; it need not be seekable. Nothing is written when the save is refused.</param>
    /// <param name="options">What to do with formulas SharpCell cannot calculate.</param>
    /// <exception cref="NotSupportedException">A cell of an Excel data table changed, or a cell holds <c>#SPILL!</c> or <c>#CALC!</c> as a value.</exception>
    /// <exception cref="InvalidOperationException">Some formulas are not calculated: call <see cref="Workbook.Recalculate"/> first; or the workbook has no sheets.</exception>
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
