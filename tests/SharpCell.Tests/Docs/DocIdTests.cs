using System.Reflection;
using SharpCell.Docs;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Docs;

public class DocIdTests
{
    [Fact]
    public void Ids_follow_the_compiler_rules_for_every_kind_of_signature()
    {
        Assert.Equal("T:SharpCell.Workbook", DocId.Of(typeof(Workbook)));
        Assert.Equal("M:SharpCell.Workbook.#ctor", DocId.Of(typeof(Workbook).GetConstructor(Type.EmptyTypes)!));
        Assert.Equal("M:SharpCell.Workbook.Evaluate(System.String,System.Threading.CancellationToken)",
            DocId.Of(typeof(Workbook).GetMethod("Evaluate", [typeof(string), typeof(CancellationToken)])!));
        Assert.Equal("M:SharpCell.Workbook.TryGetSheet(System.String,SharpCell.Worksheet@)",
            DocId.Of(typeof(Workbook).GetMethod("TryGetSheet")!));
        Assert.Equal("M:SharpCell.CellValue.Array(SharpCell.CellValue[0:,0:])",
            DocId.Of(typeof(CellValue).GetMethod("Array")!));
        Assert.Equal("M:SharpCell.ErrorKinds.TryParse(System.ReadOnlySpan{System.Char},SharpCell.ErrorKind@)",
            DocId.Of(typeof(ErrorKinds).GetMethod("TryParse")!));
        Assert.Equal("M:SharpCell.CellValue.op_Implicit(System.Boolean)~SharpCell.CellValue",
            DocId.Of(typeof(CellValue).GetMethods().Single(m => m.Name == "op_Implicit" && m.GetParameters()[0].ParameterType == typeof(bool))));
        Assert.Equal("P:SharpCell.Worksheet.Item(System.String)",
            DocId.Of(typeof(Worksheet).GetProperty("Item", [typeof(string)])!));
        Assert.Equal("P:SharpCell.Cell.Value", DocId.Of(typeof(Cell).GetProperty("Value")!));
        Assert.Equal("F:SharpCell.ErrorKind.Div0", DocId.Of(typeof(ErrorKind).GetField("Div0")!));
        Assert.Equal("M:SharpCell.CalculationDiagnostic.#ctor(SharpCell.DiagnosticKind,SharpCell.Worksheet,System.String,System.String)",
            DocId.Of(typeof(CalculationDiagnostic).GetConstructors().Single()));
        Assert.Equal("M:SharpCell.Xlsx.XlsxReader.Load(System.IO.Stream)",
            DocId.Of(typeof(XlsxReader).GetMethod("Load", [typeof(Stream)])!));
    }

    [Fact]
    public void Every_public_member_has_documentation_under_its_id()
    {
        var assemblies = ApiReference.DocumentedAssemblies();
        var docs = XmlDocs.Load(assemblies);
        var missing = assemblies
            .SelectMany(a => a.GetExportedTypes())
            .SelectMany(t => ApiReference.MembersOf(t).Prepend(t))
            .Select(DocId.Of)
            .Where(id => !docs.Has(id))
            .ToList();

        Assert.True(missing.Count == 0, "No documentation for:\n" + string.Join("\n", missing));
    }

    [Fact]
    public void Every_public_member_has_a_summary()
    {
        var assemblies = ApiReference.DocumentedAssemblies();
        var docs = XmlDocs.Load(assemblies);
        var empty = assemblies
            .SelectMany(a => a.GetExportedTypes())
            .SelectMany(t => ApiReference.MembersOf(t).Prepend(t))
            .Select(DocId.Of)
            .Where(id => docs.Summary(id, _ => null).Length == 0)
            .ToList();

        Assert.True(empty.Count == 0, "No summary for:\n" + string.Join("\n", empty));
    }
}
