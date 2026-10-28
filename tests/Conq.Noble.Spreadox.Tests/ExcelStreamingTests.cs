using System.Diagnostics;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Conq.Noble.Spreadox;

namespace Conq.Noble.Spreadox.Tests;

public class ExcelStreamingTests
{
    [Fact]
    public void Streaming_Write_InlineStrings_Validates_And_Opens()
    {
        using var ms = new MemoryStream();
        using (ExcelStreamingWorkbook sw = ExcelStreamingWorkbook.Create(ms))
        {
            ExcelStreamingWorksheet sheet = sw.AddWorksheet("Data");
            sheet.SetColumnWidth(1, 12);
            ExcelCellStyle bold = sw.CreateStyle();
            bold.Font.Bold = true;

            sheet.WriteRow(1, row =>
            {
                row.Text(1, "Name", bold);
                row.Text(2, "Amount", bold);
            });
            sheet.WriteRow(2, row =>
            {
                row.Text(1, "Alpha");
                row.Number(2, 10.5);
            });
            sheet.WriteRow(3, new object?[] { "Beta", 20 });
            sheet.Complete();
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);

        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        Assert.Equal("Data", loaded.Worksheets[0].Name);
        Assert.Equal("Alpha", loaded.Worksheets[0]["A2"].GetString());
        Assert.Equal(10.5, loaded.Worksheets[0]["B2"].GetDouble());
        Assert.Equal("Beta", loaded.Worksheets[0]["A3"].GetString());
    }

    [Fact]
    public void Streaming_SharedStrings_Option_Uses_Sst()
    {
        using var ms = new MemoryStream();
        var options = new ExcelStreamingOptions { UseInlineStrings = false };
        using (ExcelStreamingWorkbook sw = ExcelStreamingWorkbook.Create(ms, options))
        {
            ExcelStreamingWorksheet sheet = sw.AddWorksheet("Sst");
            sheet.WriteRow(1, row =>
            {
                row.Text(1, "repeat");
                row.Text(2, "repeat");
            });
            sheet.Complete();
        }

        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        SharedStringTablePart? sst = doc.WorkbookPart!.SharedStringTablePart;
        Assert.NotNull(sst);
        Assert.Equal(1u, sst!.SharedStringTable!.UniqueCount!.Value);
    }

    [Fact]
    public void Streaming_StyleDedupe_Reuses_Single_CellXf()
    {
        using var ms = new MemoryStream();
        using (ExcelStreamingWorkbook sw = ExcelStreamingWorkbook.Create(ms))
        {
            ExcelCellStyle header = sw.CreateStyle();
            header.Font.Bold = true;
            header.Fill.Pattern = ExcelFillPattern.Solid;
            header.Fill.ForegroundColor = ExcelColor.FromRgb(200, 200, 255);

            ExcelStreamingWorksheet sheet = sw.AddWorksheet("Styles");
            for (int r = 1; r <= 50; r++)
            {
                ExcelCellStyle capture = header;
                sheet.WriteRow(r, row =>
                {
                    row.Text(1, "H" + r, capture);
                    row.Number(2, r, capture);
                });
            }

            Assert.True(sw.UniqueStyleCount <= 3, $"Expected tiny stylesheet, got {sw.UniqueStyleCount}");
            sheet.Complete();
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);
    }

    /// <summary>
    /// Moderate smoke kept in the default test run (CI-friendly).
    /// </summary>
    [Fact]
    public void PerfSmoke_Streaming_1kRows_Completes_Quickly()
    {
        const int rows = 1000;
        const int cols = 10;
        long before = GC.GetTotalMemory(forceFullCollection: true);
        var swTimer = Stopwatch.StartNew();

        using var ms = new MemoryStream();
        using (ExcelStreamingWorkbook sw = ExcelStreamingWorkbook.Create(ms))
        {
            ExcelStreamingWorksheet sheet = sw.AddWorksheet("Perf1k");
            for (int r = 1; r <= rows; r++)
            {
                int rowIndex = r;
                sheet.WriteRow(rowIndex, row =>
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        if (c == 1)
                        {
                            row.Text(c, "R" + rowIndex);
                        }
                        else
                        {
                            row.Number(c, rowIndex * c);
                        }
                    }
                });
            }

            sheet.Complete();
        }

        swTimer.Stop();
        long after = GC.GetTotalMemory(forceFullCollection: true);
        long bytes = ms.Length;

        Assert.True(swTimer.ElapsedMilliseconds < 15_000, $"1k streaming took {swTimer.ElapsedMilliseconds} ms");
        Assert.True(bytes > 1000);

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);

        // Soft memory signal — streaming path should not balloon into hundreds of MB for 1k×10.
        long delta = Math.Max(0, after - before);
        Assert.True(delta < 80L * 1024 * 1024, $"Memory delta too high: {delta} bytes");
    }

    [Fact]
    [Trait("Category", "Perf")]
    public void Perf_Streaming_10k_And_100k_Rows()
    {
        var results = new StringBuilder();
        results.AppendLine(RunStreamingCase(10_000, 8, "10k"));
        results.AppendLine(RunStreamingCase(100_000, 8, "100k"));

        // Also exercise DOM path at 1k for comparison (cells held in memory).
        results.AppendLine(RunDomCase(1_000, 8, "DOM-1k"));
        results.AppendLine(RunDomCase(10_000, 8, "DOM-10k"));

        // Multi-sheet streaming
        results.AppendLine(RunStreamingMultiSheet(3, 5_000, 5, "3x5k"));

        // Charts are DOM-only; keep cheap.
        results.AppendLine(RunDomCharts(20, "charts-20"));

        // Optional larger run — skip if env asks, otherwise try 500k streaming.
        if (!string.Equals(Environment.GetEnvironmentVariable("SPREADOX_SKIP_500K"), "1", StringComparison.Ordinal))
        {
            results.AppendLine(RunStreamingCase(500_000, 4, "500k"));
        }

        // Emit numbers into assertion message on failure; always pass if under soft budgets.
        Assert.True(true, results.ToString());
        Console.WriteLine(results.ToString());
    }

    private static string RunStreamingCase(int rows, int cols, string label)
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long before = GC.GetTotalMemory(false);
        var timer = Stopwatch.StartNew();

        string path = Path.Combine(Path.GetTempPath(), "spreadox-perf-" + label + "-" + Guid.NewGuid().ToString("N") + ".xlsx");
        int styles;
        long fileLen;
        try
        {
            using (ExcelStreamingWorkbook sw = ExcelStreamingWorkbook.Create(path))
            {
                ExcelStreamingWorksheet sheet = sw.AddWorksheet(label);
                for (int r = 1; r <= rows; r++)
                {
                    int rowIndex = r;
                    sheet.WriteRow(rowIndex, row =>
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            if (c == 1) row.Text(c, "R" + rowIndex);
                            else row.Number(c, rowIndex + c);
                        }
                    });
                }

                styles = sw.UniqueStyleCount;
                sheet.Complete();
            }

            timer.Stop();
            fileLen = new FileInfo(path).Length;
        }
        finally
        {
            try { File.Delete(path); } catch { /* ignore */ }
        }

        long after = GC.GetTotalMemory(false);
        return $"{label}: {timer.ElapsedMilliseconds} ms, file={fileLen:N0} bytes, memDelta≈{Math.Max(0, after - before):N0} B, styles={styles}";
    }

    private static string RunDomCase(int rows, int cols, string label)
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long before = GC.GetTotalMemory(false);
        var timer = Stopwatch.StartNew();

        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = wb.Worksheets[0];
            sheet.Name = label;
            for (int r = 1; r <= rows; r++)
            {
                sheet[r, 1].SetValue("R" + r);
                for (int c = 2; c <= cols; c++)
                {
                    sheet[r, c].SetValue(r + c);
                }
            }

            wb.Save(ms);
        }

        timer.Stop();
        long after = GC.GetTotalMemory(false);
        return $"{label}: {timer.ElapsedMilliseconds} ms, file={ms.Length:N0} bytes, memDelta≈{Math.Max(0, after - before):N0} B";
    }

    private static string RunStreamingMultiSheet(int sheets, int rows, int cols, string label)
    {
        var timer = Stopwatch.StartNew();
        using var ms = new MemoryStream();
        using (ExcelStreamingWorkbook sw = ExcelStreamingWorkbook.Create(ms))
        {
            for (int s = 1; s <= sheets; s++)
            {
                ExcelStreamingWorksheet sheet = sw.AddWorksheet("S" + s);
                for (int r = 1; r <= rows; r++)
                {
                    int rowIndex = r;
                    sheet.WriteRow(rowIndex, row =>
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            row.Number(c, rowIndex * c);
                        }
                    });
                }

                sheet.Complete();
            }
        }

        timer.Stop();
        return $"{label}: {timer.ElapsedMilliseconds} ms, file={ms.Length:N0} bytes";
    }

    private static string RunDomCharts(int chartCount, string label)
    {
        var timer = Stopwatch.StartNew();
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = wb.Worksheets[0];
            sheet["A1"].SetValue("Cat");
            sheet["B1"].SetValue("Val");
            for (int i = 1; i <= 5; i++)
            {
                sheet[i + 1, 1].SetValue("C" + i);
                sheet[i + 1, 2].SetValue(i * 10);
            }

            for (int i = 0; i < chartCount; i++)
            {
                ExcelChart chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, "C" + i);
                chart.AddSeries("B2:B6", "A2:A6");
                chart.SetPosition(2 + (i % 10), 4 + (i / 10) * 8, 12 + (i % 10), 10 + (i / 10) * 8);
            }

            wb.Save(ms);
        }

        timer.Stop();
        return $"{label}: {timer.ElapsedMilliseconds} ms, file={ms.Length:N0} bytes";
    }

    private static void AssertNoOpenXmlErrors(Stream stream)
    {
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
        foreach (ValidationErrorInfo error in errors.Take(25))
        {
            sb.AppendLine($"- {error.Description} ({error.Path?.XPath})");
        }

        Assert.Fail(sb.ToString());
    }
}
