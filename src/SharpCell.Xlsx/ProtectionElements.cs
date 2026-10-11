using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace SharpCell.Xlsx;

/// <summary>
/// The attributes holding a protection password. The sheet's and the workbook's protection store the
/// same hash under different names (<c>hashValue</c> and <c>workbookHashValue</c>).
/// </summary>
internal sealed record PasswordAttributes(string AlgorithmName, string HashValue, string SaltValue, string SpinCount, string Legacy)
{
    public static PasswordAttributes Sheet { get; } = new("algorithmName", "hashValue", "saltValue", "spinCount", "password");

    public static PasswordAttributes Workbook { get; } = new("workbookAlgorithmName", "workbookHashValue", "workbookSaltValue", "workbookSpinCount", "workbookPassword");

    /// <summary>The names of all the attributes holding a password.</summary>
    public string[] Names => [AlgorithmName, HashValue, SaltValue, SpinCount, Legacy];

    /// <summary>The password the attributes hold, as read; null when they hold none.</summary>
    public ProtectionPassword? Read(Func<string, string?> attribute)
    {
        var password = new ProtectionPassword
        {
            AlgorithmName = attribute(AlgorithmName),
            HashValue = attribute(HashValue),
            SaltValue = attribute(SaltValue),
            SpinCount = int.TryParse(attribute(SpinCount), NumberStyles.None, CultureInfo.InvariantCulture, out var spins) ? spins : null,
            LegacyHash = attribute(Legacy),
        };
        return password.HashValue is null && password.LegacyHash is null ? null : password;
    }

    /// <summary>The attributes for a password, in the schema's order; none for no password.</summary>
    public IEnumerable<XAttribute> Write(ProtectionPassword? password)
    {
        if (password is null)
            yield break;
        if (password.LegacyHash is { } legacy)
            yield return new XAttribute(Legacy, legacy);
        if (password.AlgorithmName is { } algorithm)
            yield return new XAttribute(AlgorithmName, algorithm);
        if (password.HashValue is { } hash)
            yield return new XAttribute(HashValue, hash);
        if (password.SaltValue is { } salt)
            yield return new XAttribute(SaltValue, salt);
        if (password.SpinCount is { } spins)
            yield return new XAttribute(SpinCount, spins.ToString(CultureInfo.InvariantCulture));
    }
}

/// <summary>
/// A sheet's protection in its &lt;sheetProtection&gt; element. The element's flags state what is
/// protected, the opposite of <see cref="SheetProtection"/>: <c>formatCells="0"</c> allows formatting,
/// <c>selectLockedCells="1"</c> forbids selecting locked cells. The sheet is protected only with
/// <c>sheet="1"</c>; every attribute of the element is modelled, so a changed element is built anew.
/// </summary>
internal static class ProtectionElements
{
    // Each flag in the schema's order, with its default there and the allowed action it states the opposite of.
    private static readonly (string Name, bool Default, Func<SheetProtection, bool> Allows, Func<SheetProtection, bool, SheetProtection> With)[] Flags =
    [
        ("objects", false, p => p.EditObjects, (p, v) => new(p) { EditObjects = v }),
        ("scenarios", false, p => p.EditScenarios, (p, v) => new(p) { EditScenarios = v }),
        ("formatCells", true, p => p.FormatCells, (p, v) => new(p) { FormatCells = v }),
        ("formatColumns", true, p => p.FormatColumns, (p, v) => new(p) { FormatColumns = v }),
        ("formatRows", true, p => p.FormatRows, (p, v) => new(p) { FormatRows = v }),
        ("insertColumns", true, p => p.InsertColumns, (p, v) => new(p) { InsertColumns = v }),
        ("insertRows", true, p => p.InsertRows, (p, v) => new(p) { InsertRows = v }),
        ("insertHyperlinks", true, p => p.InsertHyperlinks, (p, v) => new(p) { InsertHyperlinks = v }),
        ("deleteColumns", true, p => p.DeleteColumns, (p, v) => new(p) { DeleteColumns = v }),
        ("deleteRows", true, p => p.DeleteRows, (p, v) => new(p) { DeleteRows = v }),
        ("selectLockedCells", false, p => p.SelectLockedCells, (p, v) => new(p) { SelectLockedCells = v }),
        ("sort", true, p => p.Sort, (p, v) => new(p) { Sort = v }),
        ("autoFilter", true, p => p.AutoFilter, (p, v) => new(p) { AutoFilter = v }),
        ("pivotTables", true, p => p.PivotTables, (p, v) => new(p) { PivotTables = v }),
        ("selectUnlockedCells", false, p => p.SelectUnlockedCells, (p, v) => new(p) { SelectUnlockedCells = v }),
    ];

    /// <summary>Sets the sheet's protection and password from the element; a sheet without <c>sheet="1"</c> stays unprotected.</summary>
    public static void Read(XElement element, Worksheet sheet)
    {
        if (!Flag(element, "sheet", false))
            return;
        var protection = new SheetProtection();
        foreach (var (name, defaultValue, _, with) in Flags)
            protection = with(protection, !Flag(element, name, defaultValue));
        sheet.SetProtection(protection, PasswordAttributes.Sheet.Read(name => (string?)element.Attribute(name)));
    }

    /// <summary>The element for the sheet's protection, flags at the schema's defaults left out; null when it is not protected.</summary>
    public static XElement? Element(XNamespace ns, Worksheet sheet)
    {
        if (sheet.Protection is not { } protection)
            return null;
        var element = new XElement(ns + "sheetProtection", PasswordAttributes.Sheet.Write(sheet.ProtectionPassword), new XAttribute("sheet", "1"));
        foreach (var (name, defaultValue, allows, _) in Flags)
        {
            var isProtected = !allows(protection);
            if (isProtected != defaultValue)
                element.Add(new XAttribute(name, isProtected ? "1" : "0"));
        }

        return element;
    }

    // An xsd:boolean attribute, or the schema's default when it is missing or not a boolean.
    private static bool Flag(XElement element, string name, bool defaultValue) => (string?)element.Attribute(name) switch
    {
        "1" or "true" => true,
        "0" or "false" => false,
        _ => defaultValue,
    };
}
