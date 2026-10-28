using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Conq.Noble.Spreadox;

namespace Conq.Noble.Spreadox.Tests;

public class ExcelWorkbookTests
{
    [Fact]
    public void Create_HasDefaultSheet()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Assert.Single(workbook.Worksheets);
        Assert.Equal("Sheet1", workbook.Worksheets[0].Name);
    }

    [Fact]
    public void AddRemoveRename_Worksheets()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        ExcelWorksheet data = workbook.AddWorksheet("Data");
        workbook.AddWorksheet("Summary");

        Assert.Equal(3, workbook.Worksheets.Count);

        workbook.RenameWorksheet("Data", "Facts");
        Assert.Equal("Facts", data.Name);
        Assert.Same(data, workbook["Facts"]);

        Assert.True(workbook.RemoveWorksheet("Sheet1"));
        Assert.Equal(2, workbook.Worksheets.Count);
        Assert.Throws<InvalidOperationException>(() =>
        {
            workbook.RemoveWorksheet("Facts");
            workbook.RemoveWorksheet("Summary");
        });
    }

    [Fact]
    public void DuplicateWorksheetName_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Assert.Throws<InvalidOperationException>(() => workbook.AddWorksheet("Sheet1"));
    }

    [Fact]
    public void CellTypes_RoundTrip_FileAndValidate()
    {
        string path = Path.Combine(Path.GetTempPath(), "Spreadox_" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            DateTime stamp = new DateTime(2024, 6, 15, 13, 45, 30, DateTimeKind.Unspecified);

            using (ExcelWorkbook workbook = ExcelWorkbook.Create())
            {
                ExcelWorksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Types";
                sheet["A1"].SetValue("hello");
                sheet["B1"].SetValue(42);
                sheet["C1"].SetValue(3.14159);
                sheet["D1"].SetValue(123.456789m);
                sheet["E1"].SetValue(true);
                sheet["F1"].SetValue(stamp);
                sheet["G1"].Clear(); // null / empty
                sheet.SetValue("H1", null);

                workbook.Save(path);
            }

            Assert.True(File.Exists(path));
            AssertNoOpenXmlErrors(path);

            using ExcelWorkbook loaded = ExcelWorkbook.Open(path);
            ExcelWorksheet types = loaded["Types"];

            Assert.Equal("hello", types["A1"].GetString());
            Assert.Equal(CellValueKind.String, types["A1"].Kind);

            Assert.Equal(42, types["B1"].GetInt());
            Assert.Equal(CellValueKind.Integer, types["B1"].Kind);

            Assert.Equal(3.14159, types["C1"].GetDouble()!.Value, 5);
            Assert.Equal(CellValueKind.Double, types["C1"].Kind);

            // OOXML numeric cells are doubles; decimal in-memory values round-trip as double.
            Assert.Equal(CellValueKind.Double, types["D1"].Kind);
            Assert.Equal(123.456789, types["D1"].GetDouble()!.Value, 5);

            Assert.True(types["E1"].GetBoolean());
            Assert.Equal(CellValueKind.Boolean, types["E1"].Kind);

            Assert.Equal(stamp, types["F1"].GetDateTime());
            Assert.Equal(CellValueKind.DateTime, types["F1"].Kind);

            Assert.True(types["G1"].IsEmpty);
            Assert.True(types["H1"].IsEmpty);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void SaveAndOpen_Stream_RoundTrip()
    {
        using var memory = new MemoryStream();
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.AddWorksheet("StreamSheet");
            workbook["StreamSheet"]["A1"].SetValue("stream-ok");
            workbook["StreamSheet"]["A2"].SetValue(99);
            workbook.Save(memory);
        }

        Assert.True(memory.Length > 0);
        memory.Position = 0;
        AssertNoOpenXmlErrors(memory);

        memory.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(memory);
        Assert.Equal("stream-ok", loaded["StreamSheet"]["A1"].GetString());
        Assert.Equal(99, loaded["StreamSheet"]["A2"].GetInt());
    }

    [Fact]
    public void Create_Save_PassesOpenXmlValidator()
    {
        using var memory = new MemoryStream();
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = workbook.AddWorksheet("Validator");
            sheet["A1"].SetValue("x");
            sheet["A2"].SetValue(1);
            sheet["A3"].SetValue(2.5);
            sheet["A4"].SetValue(false);
            sheet["A5"].SetValue(DateTime.UtcNow.Date);
            workbook.Save(memory);
        }

        memory.Position = 0;
        AssertNoOpenXmlErrors(memory);
    }

    [Fact]
    public void Dispose_PreventsFurtherUse()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        ExcelWorksheet sheet = workbook.Worksheets[0];
        workbook.Dispose();
        Assert.Throws<ObjectDisposedException>(() => workbook.AddWorksheet("X"));
        Assert.Throws<ObjectDisposedException>(() => workbook.Save(new MemoryStream()));
        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets);
        Assert.Throws<ObjectDisposedException>(() => _ = sheet["A1"]);
        Assert.Throws<ObjectDisposedException>(() => sheet.Charts.Add(ExcelChartType.Line));
    }

    [Fact]
    public void Open_RejectsNullOrNonReadableStream()
    {
        Assert.Throws<ArgumentNullException>(() => ExcelWorkbook.Open((Stream)null!));
        using var ms = new MemoryStream();
        ms.Dispose();
        Assert.ThrowsAny<Exception>(() => ExcelWorkbook.Open(ms));
    }

    [Fact]
    public void PublicApi_DoesNotExposeOpenXmlNamespaces_OnCoreTypes()
    {
        // Smoke: consumers only need Conq.Noble.Spreadox types for Phase 1 workflows.
        Type[] publicTypes =
        {
            typeof(ExcelWorkbook),
            typeof(ExcelWorksheet),
            typeof(ExcelCell),
            typeof(CellValueKind),
            typeof(CellAddress)
        };

        foreach (Type type in publicTypes)
        {
            Assert.Equal("Conq.Noble.Spreadox", type.Namespace);
            Assert.DoesNotContain("DocumentFormat", type.Assembly.GetName().Name!);
        }

        Assembly assembly = typeof(ExcelWorkbook).Assembly;
        Assert.Equal("Conq.Noble.Spreadox", assembly.GetName().Name);
        Assert.Equal("Conq.Noble", assembly.GetCustomAttribute<AssemblyCompanyAttribute>()?.Company);
        Assert.Equal("Spreadox", assembly.GetCustomAttribute<AssemblyProductAttribute>()?.Product);
        string copyright = assembly.GetCustomAttribute<AssemblyCopyrightAttribute>()?.Copyright ?? string.Empty;
        Assert.Contains("Conq.Noble", copyright);
    }

    private static void AssertNoOpenXmlErrors(string path)
    {
        using FileStream stream = File.OpenRead(path);
        AssertNoOpenXmlErrors(stream);
    }

    private static void AssertNoOpenXmlErrors(Stream stream)
    {
        // Copy: SpreadsheetDocument.Dispose closes the stream it opens.
        using MemoryStream copy = new MemoryStream();
        long original = stream.CanSeek ? stream.Position : 0;
        stream.CopyTo(copy);
        if (stream.CanSeek)
        {
            stream.Position = original;
        }

        copy.Position = 0;
        using SpreadsheetDocument document = SpreadsheetDocument.Open(copy, false);
        OpenXmlValidator validator = new OpenXmlValidator();
        List<ValidationErrorInfo> errors = validator.Validate(document).ToList();
        if (errors.Count == 0)
        {
            return;
        }

        var sb = new StringBuilder();
        sb.AppendLine($"OpenXmlValidator reported {errors.Count} error(s):");
        foreach (ValidationErrorInfo error in errors.Take(20))
        {
            sb.AppendLine($"- {error.Description} ({error.Path?.XPath})");
        }

        Assert.True(false, sb.ToString());
    }
}

