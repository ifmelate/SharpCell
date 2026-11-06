using System;
using System.Text;

namespace SharpCell;

/// <summary>
/// The paper a sheet prints on, by the number Excel stores (as its <c>XlPaperSize</c> does). A file
/// may hold a number not named here, from a printer's own sizes; it is kept as it is.
/// </summary>
public enum PaperSize
{
    /// <summary>Letter, 8.5 by 11 inches; what a sheet prints on unless set otherwise.</summary>
    Letter = 1,

    /// <summary>Letter Small, 8.5 by 11 inches.</summary>
    LetterSmall = 2,

    /// <summary>Tabloid, 11 by 17 inches.</summary>
    Tabloid = 3,

    /// <summary>Ledger, 17 by 11 inches.</summary>
    Ledger = 4,

    /// <summary>Legal, 8.5 by 14 inches.</summary>
    Legal = 5,

    /// <summary>Statement, 5.5 by 8.5 inches.</summary>
    Statement = 6,

    /// <summary>Executive, 7.25 by 10.5 inches.</summary>
    Executive = 7,

    /// <summary>A3, 297 by 420 mm.</summary>
    A3 = 8,

    /// <summary>A4, 210 by 297 mm.</summary>
    A4 = 9,

    /// <summary>A4 Small, 210 by 297 mm.</summary>
    A4Small = 10,

    /// <summary>A5, 148 by 210 mm.</summary>
    A5 = 11,

    /// <summary>B4 (JIS), 257 by 364 mm.</summary>
    B4 = 12,

    /// <summary>B5 (JIS), 182 by 257 mm.</summary>
    B5 = 13,

    /// <summary>Folio, 8.5 by 13 inches.</summary>
    Folio = 14,

    /// <summary>Quarto, 215 by 275 mm.</summary>
    Quarto = 15,

    /// <summary>10 by 14 inches.</summary>
    Size10x14 = 16,

    /// <summary>11 by 17 inches.</summary>
    Size11x17 = 17,

    /// <summary>Note, 8.5 by 11 inches.</summary>
    Note = 18,

    /// <summary>Envelope #9, 3.875 by 8.875 inches.</summary>
    Envelope9 = 19,

    /// <summary>Envelope #10, 4.125 by 9.5 inches.</summary>
    Envelope10 = 20,

    /// <summary>Envelope #11, 4.5 by 10.375 inches.</summary>
    Envelope11 = 21,

    /// <summary>Envelope #12, 4.75 by 11 inches.</summary>
    Envelope12 = 22,

    /// <summary>Envelope #14, 5 by 11.5 inches.</summary>
    Envelope14 = 23,

    /// <summary>C size sheet, 17 by 22 inches.</summary>
    CSheet = 24,

    /// <summary>D size sheet, 22 by 34 inches.</summary>
    DSheet = 25,

    /// <summary>E size sheet, 34 by 44 inches.</summary>
    ESheet = 26,

    /// <summary>Envelope DL, 110 by 220 mm.</summary>
    EnvelopeDL = 27,

    /// <summary>Envelope C5, 162 by 229 mm.</summary>
    EnvelopeC5 = 28,

    /// <summary>Envelope C3, 324 by 458 mm.</summary>
    EnvelopeC3 = 29,

    /// <summary>Envelope C4, 229 by 324 mm.</summary>
    EnvelopeC4 = 30,

    /// <summary>Envelope C6, 114 by 162 mm.</summary>
    EnvelopeC6 = 31,

    /// <summary>Envelope C65, 114 by 229 mm.</summary>
    EnvelopeC65 = 32,

    /// <summary>Envelope B4, 250 by 353 mm.</summary>
    EnvelopeB4 = 33,

    /// <summary>Envelope B5, 176 by 250 mm.</summary>
    EnvelopeB5 = 34,

    /// <summary>Envelope B6, 176 by 125 mm.</summary>
    EnvelopeB6 = 35,

    /// <summary>Envelope Italy, 110 by 230 mm.</summary>
    EnvelopeItaly = 36,

    /// <summary>Envelope Monarch, 3.875 by 7.5 inches.</summary>
    EnvelopeMonarch = 37,

    /// <summary>Envelope 6 3/4, 3.625 by 6.5 inches.</summary>
    EnvelopePersonal = 38,

    /// <summary>US Standard Fanfold, 14.875 by 11 inches.</summary>
    FanfoldUS = 39,

    /// <summary>German Standard Fanfold, 8.5 by 12 inches.</summary>
    FanfoldStandardGerman = 40,

    /// <summary>German Legal Fanfold, 8.5 by 13 inches.</summary>
    FanfoldLegalGerman = 41,

    /// <summary>A2, 420 by 594 mm.</summary>
    A2 = 66,

    /// <summary>A6, 105 by 148 mm.</summary>
    A6 = 70,
}

/// <summary>
/// How many pages a printed sheet is shrunk to fit, as Excel's Page Setup &gt; Fit to does:
/// so many pages wide by so many tall. 0 leaves that direction to as many pages as it takes, as
/// a blank box does in Excel; <c>new PageFit(1, 0)</c> fits all columns on one page wide.
/// </summary>
public readonly record struct PageFit
{
    private const int Max = 32767;

    /// <summary>A fit of so many pages wide by so many tall.</summary>
    /// <param name="wide">Pages across, 0 for as many as it takes.</param>
    /// <param name="tall">Pages down, 0 for as many as it takes.</param>
    /// <exception cref="ArgumentOutOfRangeException">A count is negative or above 32767.</exception>
    public PageFit(int wide, int tall)
    {
        Wide = Check(wide, nameof(wide));
        Tall = Check(tall, nameof(tall));
    }

    /// <summary>Pages across; 0 for as many as it takes.</summary>
    public int Wide { get; }

    /// <summary>Pages down; 0 for as many as it takes.</summary>
    public int Tall { get; }

    private static int Check(int value, string name) =>
        value is >= 0 and <= Max ? value : throw new ArgumentOutOfRangeException(name, value, $"A page count is from 0 to {Max}.");
}

/// <summary>
/// A printed header or footer: its left, centre and right sections, in Excel's codes. Text prints
/// as written, apart from codes such as <see cref="PageNumber"/> (<c>&amp;P</c>), <c>&amp;B</c> for
/// bold or <c>&amp;"Arial,Italic"</c> for a font; <c>&amp;&amp;</c> prints one ampersand, and
/// <see cref="Escape"/> doubles those of plain text. Null or empty leaves a section blank.
/// </summary>
/// <param name="Left">The section on the left, or null.</param>
/// <param name="Center">The section in the middle, or null.</param>
/// <param name="Right">The section on the right, or null.</param>
public sealed record HeaderFooter(string? Left = null, string? Center = null, string? Right = null)
{
    /// <summary>The code for the page number.</summary>
    public static string PageNumber => "&P";

    /// <summary>The code for the number of pages.</summary>
    public static string PageCount => "&N";

    /// <summary>The code for the date of printing.</summary>
    public static string Date => "&D";

    /// <summary>The code for the time of printing.</summary>
    public static string Time => "&T";

    /// <summary>The code for the file's name.</summary>
    public static string FileName => "&F";

    /// <summary>The code for the file's folder and name.</summary>
    public static string FilePath => "&Z&F";

    /// <summary>The code for the sheet's name.</summary>
    public static string SheetName => "&A";

    /// <summary>Plain text with its ampersands doubled, so that it prints as written.</summary>
    /// <exception cref="ArgumentNullException">The text is null.</exception>
    public static string Escape(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        return text.Replace("&", "&&", StringComparison.Ordinal);
    }

    /// <summary>
    /// Splits a header or footer as an .xlsx file stores it, such as <c>&amp;LLeft&amp;CMiddle</c>,
    /// into its sections. Text before any section code belongs to the middle, as in Excel.
    /// </summary>
    /// <exception cref="ArgumentNullException">The text is null.</exception>
    public static HeaderFooter Parse(string code)
    {
        ArgumentNullException.ThrowIfNull(code);
        var sections = new StringBuilder?[3];
        var current = 1;
        var start = 0;
        for (var i = 0; i < code.Length; i++)
        {
            if (code[i] != '&' || i + 1 == code.Length)
                continue;
            var next = code[i + 1];
            if (next == '"')
            {
                // A font name runs to the next quote, and may hold anything.
                var close = code.IndexOf('"', i + 2);
                i = close < 0 ? code.Length : close;
                continue;
            }

            var section = next switch { 'L' => 0, 'C' => 1, 'R' => 2, _ => -1 };
            if (section < 0)
            {
                i++;   // && and the other codes stay as they are
                continue;
            }

            Append(sections, current, code, start, i);
            current = section;
            start = i + 2;
            i++;
        }

        Append(sections, current, code, start, code.Length);
        return new HeaderFooter(sections[0]?.ToString(), sections[1]?.ToString(), sections[2]?.ToString());
    }

    /// <summary>The header or footer as an .xlsx file stores it, such as <c>&amp;LLeft&amp;CMiddle</c>; empty when every section is blank.</summary>
    public override string ToString()
    {
        var text = new StringBuilder();
        if (!string.IsNullOrEmpty(Left))
            text.Append("&L").Append(Left);
        if (!string.IsNullOrEmpty(Center))
            text.Append("&C").Append(Center);
        if (!string.IsNullOrEmpty(Right))
            text.Append("&R").Append(Right);
        return text.ToString();
    }

    internal bool IsBlank => string.IsNullOrEmpty(Left) && string.IsNullOrEmpty(Center) && string.IsNullOrEmpty(Right);

    private static void Append(StringBuilder?[] sections, int section, string code, int start, int end)
    {
        if (end > start)
            (sections[section] ??= new StringBuilder()).Append(code, start, end - start);
    }
}
