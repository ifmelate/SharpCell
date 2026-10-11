using System;
using System.Collections.Generic;

namespace SharpCell;

/// <summary>
/// What users may still do on a protected sheet, as Excel's Protect Sheet dialog lists it: by default
/// they may select cells and nothing else. Set properties in an initializer and pass the object to
/// <see cref="Worksheet.Protect"/>. Excel enforces the protection; SharpCell only keeps and saves it.
/// </summary>
public sealed class SheetProtection : IEquatable<SheetProtection>
{
    /// <summary>A protection allowing users to select cells and nothing else; set properties in an initializer.</summary>
    public SheetProtection()
    {
    }

    /// <summary>
    /// A copy of another protection, so it can be changed in an initializer:
    /// <c>new SheetProtection(sheet.Protection) { Sort = true }</c>.
    /// </summary>
    /// <param name="source">The protection to copy.</param>
    /// <exception cref="ArgumentNullException">The source is null.</exception>
    public SheetProtection(SheetProtection source)
    {
        ArgumentNullException.ThrowIfNull(source);
        SelectLockedCells = source.SelectLockedCells;
        SelectUnlockedCells = source.SelectUnlockedCells;
        FormatCells = source.FormatCells;
        FormatColumns = source.FormatColumns;
        FormatRows = source.FormatRows;
        InsertColumns = source.InsertColumns;
        InsertRows = source.InsertRows;
        InsertHyperlinks = source.InsertHyperlinks;
        DeleteColumns = source.DeleteColumns;
        DeleteRows = source.DeleteRows;
        Sort = source.Sort;
        AutoFilter = source.AutoFilter;
        PivotTables = source.PivotTables;
        EditObjects = source.EditObjects;
        EditScenarios = source.EditScenarios;
    }

    /// <summary>Whether users may select locked cells; true unless set.</summary>
    public bool SelectLockedCells { get; init; } = true;

    /// <summary>Whether users may select unlocked cells; true unless set.</summary>
    public bool SelectUnlockedCells { get; init; } = true;

    /// <summary>Whether users may format cells; false unless set.</summary>
    public bool FormatCells { get; init; }

    /// <summary>Whether users may format columns, their widths and hiding included; false unless set.</summary>
    public bool FormatColumns { get; init; }

    /// <summary>Whether users may format rows, their heights and hiding included; false unless set.</summary>
    public bool FormatRows { get; init; }

    /// <summary>Whether users may insert columns; false unless set.</summary>
    public bool InsertColumns { get; init; }

    /// <summary>Whether users may insert rows; false unless set.</summary>
    public bool InsertRows { get; init; }

    /// <summary>Whether users may insert hyperlinks; false unless set.</summary>
    public bool InsertHyperlinks { get; init; }

    /// <summary>Whether users may delete columns; false unless set.</summary>
    public bool DeleteColumns { get; init; }

    /// <summary>Whether users may delete rows; false unless set.</summary>
    public bool DeleteRows { get; init; }

    /// <summary>Whether users may sort; false unless set.</summary>
    public bool Sort { get; init; }

    /// <summary>Whether users may use AutoFilter; false unless set.</summary>
    public bool AutoFilter { get; init; }

    /// <summary>Whether users may use PivotTables and PivotCharts; false unless set.</summary>
    public bool PivotTables { get; init; }

    /// <summary>Whether users may edit objects such as shapes and charts; false unless set.</summary>
    public bool EditObjects { get; init; }

    /// <summary>Whether users may edit scenarios; false unless set.</summary>
    public bool EditScenarios { get; init; }

    /// <summary>Whether the other protection allows the same actions.</summary>
    public bool Equals(SheetProtection? other) => other is not null && Flags == other.Flags;

    /// <summary>Whether <paramref name="obj"/> is a protection allowing the same actions.</summary>
    public override bool Equals(object? obj) => Equals(obj as SheetProtection);

    /// <summary>A hash code consistent with <see cref="Equals(SheetProtection)"/>.</summary>
    public override int GetHashCode() => Flags;

    /// <summary>The names of the allowed actions, such as <c>SelectLockedCells, SelectUnlockedCells</c>, or <c>None</c>.</summary>
    public override string ToString()
    {
        var allowed = new List<string>();
        foreach (var (name, allows) in Actions)
        {
            if (allows)
                allowed.Add(name);
        }

        return allowed.Count == 0 ? "None" : string.Join(", ", allowed);
    }

    private (string Name, bool Allows)[] Actions =>
    [
        (nameof(SelectLockedCells), SelectLockedCells), (nameof(SelectUnlockedCells), SelectUnlockedCells),
        (nameof(FormatCells), FormatCells), (nameof(FormatColumns), FormatColumns), (nameof(FormatRows), FormatRows),
        (nameof(InsertColumns), InsertColumns), (nameof(InsertRows), InsertRows), (nameof(InsertHyperlinks), InsertHyperlinks),
        (nameof(DeleteColumns), DeleteColumns), (nameof(DeleteRows), DeleteRows), (nameof(Sort), Sort),
        (nameof(AutoFilter), AutoFilter), (nameof(PivotTables), PivotTables), (nameof(EditObjects), EditObjects),
        (nameof(EditScenarios), EditScenarios),
    ];

    // One bit per action, for equality and hashing.
    private int Flags
    {
        get
        {
            var flags = 0;
            foreach (var (_, allows) in Actions)
                flags = (flags << 1) | (allows ? 1 : 0);
            return flags;
        }
    }
}
