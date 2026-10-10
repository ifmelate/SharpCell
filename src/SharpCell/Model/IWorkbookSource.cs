using System;

namespace SharpCell;

/// <summary>What a file reader keeps in <see cref="Workbook.Source"/>: it must follow the workbook into a clone.</summary>
internal interface IWorkbookSource
{
    /// <summary>The same source for a copy of the workbook, whose sheets <paramref name="map"/> gives. Must only read this source.</summary>
    IWorkbookSource CopyFor(Func<Worksheet, Worksheet> map);
}
