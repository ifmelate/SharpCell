using System;
using System.Collections.Generic;

namespace SharpCell.Parsing;

/// <summary>
/// Functions added to Excel after the 2007 file format: files store their calls with a
/// <c>_xlfn.</c> prefix (FILTER and SORT with <c>_xlfn._xlws.</c>), and Excel reads a call without
/// it as an unknown function, <c>#NAME?</c>.
/// </summary>
internal static class FutureFunctions
{
    private static readonly HashSet<string> Names = new(StringComparer.Ordinal)
    {
        // Excel 2010
        "AGGREGATE", "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.INV", "CEILING.PRECISE", "CHISQ.DIST",
        "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST", "CONFIDENCE.NORM", "CONFIDENCE.T",
        "COVARIANCE.P", "COVARIANCE.S", "ERF.PRECISE", "ERFC.PRECISE", "EXPON.DIST", "F.DIST", "F.DIST.RT",
        "F.INV", "F.INV.RT", "F.TEST", "FLOOR.PRECISE", "GAMMA.DIST", "GAMMA.INV", "GAMMALN.PRECISE",
        "HYPGEOM.DIST", "LOGNORM.DIST", "LOGNORM.INV", "MODE.MULT", "MODE.SNGL", "NEGBINOM.DIST",
        "NETWORKDAYS.INTL", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "PERCENTILE.EXC",
        "PERCENTILE.INC", "PERCENTRANK.EXC", "PERCENTRANK.INC", "POISSON.DIST", "QUARTILE.EXC",
        "QUARTILE.INC", "RANK.AVG", "RANK.EQ", "STDEV.P", "STDEV.S", "T.DIST", "T.DIST.2T", "T.DIST.RT",
        "T.INV", "T.INV.2T", "T.TEST", "VAR.P", "VAR.S", "WEIBULL.DIST", "WORKDAY.INTL", "Z.TEST",

        // Excel 2013
        "ACOT", "ACOTH", "ARABIC", "BASE", "BINOM.DIST.RANGE", "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT",
        "BITXOR", "CEILING.MATH", "COMBINA", "COT", "COTH", "CSC", "CSCH", "DAYS", "DECIMAL", "ECMA.CEILING",
        "ENCODEURL", "FILTERXML", "FLOOR.MATH", "FORMULATEXT", "GAMMA", "GAUSS", "IFNA", "IMCOSH", "IMCOT",
        "IMCSC", "IMCSCH", "IMSEC", "IMSECH", "IMSINH", "IMTAN", "ISFORMULA", "ISOWEEKNUM", "MUNIT",
        "NUMBERVALUE", "PDURATION", "PERMUTATIONA", "PHI", "QUERYSTRING", "RRI", "SEC", "SECH", "SHEET",
        "SHEETS", "SKEW.P", "UNICHAR", "UNICODE", "WEBSERVICE", "XOR",

        // Excel 2016 and 2019
        "CONCAT", "FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY", "FORECAST.ETS.STAT",
        "FORECAST.LINEAR", "IFS", "MAXIFS", "MINIFS", "SWITCH", "TEXTJOIN",

        // Microsoft 365: dynamic arrays, LAMBDA and later additions
        "ANCHORARRAY", "ARRAYTOTEXT", "BYCOL", "BYROW", "CHOOSECOLS", "CHOOSEROWS", "DROP", "EXPAND", "FILTER",
        "HSTACK", "ISOMITTED", "LAMBDA", "LET", "MAKEARRAY", "MAP", "RANDARRAY", "REDUCE", "SCAN", "SEQUENCE",
        "SINGLE", "SORT", "SORTBY", "TAKE", "TEXTAFTER", "TEXTBEFORE", "TEXTSPLIT", "TOCOL", "TOROW", "UNIQUE",
        "VALUETOTEXT", "VSTACK", "WRAPCOLS", "WRAPROWS", "XLOOKUP", "XMATCH", "REGEXTEST", "REGEXEXTRACT",
        "REGEXREPLACE", "GROUPBY", "PIVOTBY", "PERCENTOF", "TRIMRANGE", "IMAGE", "STOCKHISTORY", "FIELDVALUE",
        "DETECTLANGUAGE", "TRANSLATE", "COPILOT",
    };

    /// <summary>The prefix a call of the function needs in a file: <c>_xlfn.</c>, <c>_xlfn._xlws.</c> or none.</summary>
    /// <param name="upperName">The function name in upper case, without prefixes.</param>
    public static string Prefix(string upperName) =>
        upperName is "FILTER" or "SORT" ? "_xlfn._xlws."
        : Names.Contains(upperName) ? "_xlfn."
        : "";
}
