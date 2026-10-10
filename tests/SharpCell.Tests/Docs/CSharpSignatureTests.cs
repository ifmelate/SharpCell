using SharpCell.Docs;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Docs;

public class CSharpSignatureTests
{
    [Fact]
    public void Type_declarations_name_the_kind()
    {
        Assert.Equal("public sealed class Workbook", CSharpSignature.Declaration(typeof(Workbook)));
        Assert.Equal("public readonly struct CellValue", CSharpSignature.Declaration(typeof(CellValue)));
        Assert.Equal("public enum ErrorKind", CSharpSignature.Declaration(typeof(ErrorKind)));
        Assert.Equal("public static class XlsxReader", CSharpSignature.Declaration(typeof(XlsxReader)));
        Assert.Equal("public sealed class FormulaParseException : Exception", CSharpSignature.Declaration(typeof(FormulaParseException)));
    }

    [Fact]
    public void Member_signatures_read_like_csharp_with_nullability()
    {
        Assert.Equal("public Workbook()", CSharpSignature.Of(typeof(Workbook).GetConstructor(Type.EmptyTypes)!));
        Assert.Equal("public CellValue Evaluate(string formula, CancellationToken cancellationToken)",
            CSharpSignature.Of(typeof(Workbook).GetMethod("Evaluate", [typeof(string), typeof(CancellationToken)])!));
        Assert.Equal("public RecalculationResult Recalculate(CancellationToken cancellationToken = default)",
            CSharpSignature.Of(typeof(Workbook).GetMethod("Recalculate")!));
        Assert.Equal("public bool TryGetSheet(string name, out Worksheet? sheet)",
            CSharpSignature.Of(typeof(Workbook).GetMethod("TryGetSheet")!));
        Assert.Equal("public void DefineName(string name, string formula, Worksheet? scope = null)",
            CSharpSignature.Of(typeof(Workbook).GetMethod("DefineName")!));
        Assert.Equal("public IReadOnlyList<Worksheet> Sheets { get; }", CSharpSignature.Of(typeof(Workbook).GetProperty("Sheets")!));
        Assert.Equal("public string? Formula { get; set; }", CSharpSignature.Of(typeof(Cell).GetProperty("Formula")!));
        Assert.Equal("public Cell this[string address] { get; }", CSharpSignature.Of(typeof(Worksheet).GetProperty("Item", [typeof(string)])!));
        Assert.Equal("public static CellValue Array(CellValue[,] values)", CSharpSignature.Of(typeof(CellValue).GetMethod("Array")!));
        Assert.Equal("public static string ToText(this ErrorKind kind)", CSharpSignature.Of(typeof(ErrorKinds).GetMethod("ToText")!));
        Assert.Equal("public static bool TryParse(ReadOnlySpan<char> text, out ErrorKind kind)",
            CSharpSignature.Of(typeof(ErrorKinds).GetMethod("TryParse")!));
        Assert.Equal("public static implicit operator CellValue(bool value)",
            CSharpSignature.Of(typeof(CellValue).GetMethods().Single(m => m.Name == "op_Implicit" && m.GetParameters()[0].ParameterType == typeof(bool))));
        Assert.Equal("public static bool operator ==(CellValue left, CellValue right)",
            CSharpSignature.Of(typeof(CellValue).GetMethod("op_Equality")!));
        Assert.Equal("public override string ToString()", CSharpSignature.Of(typeof(Cell).GetMethod("ToString", Type.EmptyTypes)!));
        Assert.Equal("public static CellValue Empty { get; }", CSharpSignature.Of(typeof(CellValue).GetProperty("Empty")!));
        Assert.Equal("Div0 = 2", CSharpSignature.Of(typeof(ErrorKind).GetField("Div0")!));
    }

    [Fact]
    public void Signatures_can_be_written_from_several_threads_at_once()
    {
        var members = ApiReference.DocumentedAssemblies().SelectMany(a => a.GetExportedTypes()).SelectMany(ApiReference.MembersOf).ToArray();
        // Parallel first: a shared cache inside the writer fails while it is being filled, not after.
        var actual = new string[members.Length];
        Parallel.For(0, members.Length, new ParallelOptions { MaxDegreeOfParallelism = 8 }, i => actual[i] = CSharpSignature.Of(members[i]));
        Assert.Equal(members.Select(CSharpSignature.Of), actual);
    }

    [Fact]
    public void Headings_and_anchors_tell_overloads_apart()
    {
        var evaluate = typeof(Workbook).GetMethod("Evaluate", [typeof(string)])!;
        var evaluateWithToken = typeof(Workbook).GetMethod("Evaluate", [typeof(string), typeof(CancellationToken)])!;
        Assert.Equal("Evaluate(string)", CSharpSignature.Heading(evaluate));
        Assert.Equal("Evaluate(string, CancellationToken)", CSharpSignature.Heading(evaluateWithToken));
        Assert.Equal("evaluate-string", CSharpSignature.Anchor(evaluate));
        Assert.Equal("evaluate-string-cancellationtoken", CSharpSignature.Anchor(evaluateWithToken));
        Assert.Equal("this[string]", CSharpSignature.Heading(typeof(Worksheet).GetProperty("Item", [typeof(string)])!));
        Assert.Equal("operator ==(CellValue, CellValue)", CSharpSignature.Heading(typeof(CellValue).GetMethod("op_Equality")!));
        Assert.NotEqual(CSharpSignature.Anchor(typeof(CellValue).GetMethod("op_Equality")!),
            CSharpSignature.Anchor(typeof(CellValue).GetMethod("op_Inequality")!));
        Assert.Equal("Workbook()", CSharpSignature.Heading(typeof(Workbook).GetConstructor(Type.EmptyTypes)!));
    }
}
