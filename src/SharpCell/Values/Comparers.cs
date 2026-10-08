using System;
using System.Globalization;

namespace SharpCell;

/// <summary>Number comparison that treats values differing only in floating-point noise as equal.</summary>
internal static class NumberComparer
{
    // Relative tolerance of 2^-48 (about 3.6e-15), close to the last of Excel's 15 significant digits.
    private const double Tolerance = 1.0 / (1L << 48);

    public static bool AreEqual(double a, double b)
    {
        if (a == b)
            return true;
        if (a == 0 || b == 0)
            return false;

        var difference = Math.Abs(a - b);
        return difference < Math.Abs(a) * Tolerance && difference < Math.Abs(b) * Tolerance;
    }

    public static int Compare(double a, double b) => AreEqual(a, b) ? 0 : a < b ? -1 : 1;
}

/// <summary>Text comparison for <c>=</c>, <c>&lt;</c> and friends: case-insensitive, in the workbook culture.</summary>
internal static class TextComparer
{
    public static bool AreEqual(string a, string b, CultureInfo culture) => Compare(a, b, culture) == 0;

    public static int Compare(string a, string b, CultureInfo culture) =>
        Math.Sign(culture.CompareInfo.Compare(a, b, CompareOptions.IgnoreCase));
}
