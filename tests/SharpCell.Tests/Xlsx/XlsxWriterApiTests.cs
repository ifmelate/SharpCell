using System;
using System.IO;
using SharpCell.Xlsx;

namespace SharpCell.Tests.Xlsx;

public class XlsxWriterApiTests
{
    [Fact]
    public void A_workbook_without_changes_is_saved_as_the_original_bytes()
    {
        var bytes = XlsxSourceTests.Template().Build().ToArray();
        var workbook = XlsxReader.Load(new MemoryStream(bytes));
        workbook.Recalculate();
        var output = new MemoryStream();

        XlsxWriter.Save(workbook, output);

        Assert.Equal(bytes, output.ToArray());
        Assert.True(output.CanWrite);   // left open
    }

    [Fact]
    public void Saving_to_a_path_leaves_no_temporary_file()
    {
        var folder = Directory.CreateTempSubdirectory().FullName;
        try
        {
            var workbook = XlsxReader.Load(XlsxSourceTests.Template().Build());
            workbook.Recalculate();
            var path = Path.Combine(folder, "out.xlsx");

            XlsxWriter.Save(workbook, path);

            Assert.Equal([path], Directory.GetFiles(folder));
            Assert.Equal(6, XlsxReader.Load(path)["S"]["A2"].Value.AsNumber());
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void A_refused_save_writes_nothing()
    {
        var folder = Directory.CreateTempSubdirectory().FullName;
        try
        {
            var workbook = XlsxReader.Load(XlsxSourceTests.Template().Build());   // not recalculated
            var output = new MemoryStream();
            var path = Path.Combine(folder, "out.xlsx");

            Assert.Throws<InvalidOperationException>(() => XlsxWriter.Save(workbook, output));
            Assert.Throws<InvalidOperationException>(() => XlsxWriter.Save(workbook, path));

            Assert.Equal(0, output.Length);
            Assert.Empty(Directory.GetFiles(folder));
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void Arguments_are_checked()
    {
        var workbook = new Workbook();
        Assert.Throws<ArgumentNullException>(() => XlsxWriter.Save(null!, new MemoryStream()));
        Assert.Throws<ArgumentNullException>(() => XlsxWriter.Save(workbook, (Stream)null!));
        Assert.Throws<ArgumentNullException>(() => XlsxWriter.Save(workbook, (string)null!));
        Assert.Throws<ArgumentNullException>(() => XlsxWriter.Save(workbook, new MemoryStream(), null!));
    }
}
