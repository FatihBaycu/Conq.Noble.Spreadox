using System.Text;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Conq.Noble.Spreadox;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Conq.Noble.Spreadox.Tests;

public class ExcelChartTests
{
    [Theory]
    [InlineData(ExcelChartType.ColumnClustered)]
    [InlineData(ExcelChartType.BarClustered)]
    [InlineData(ExcelChartType.Line)]
    [InlineData(ExcelChartType.Scatter)]
    public void ChartTypes_Write_HaveChartAndDrawingParts(ExcelChartType type)
    {
        using MemoryStream ms = CreateWorkbookWithChart(type, multiSeries: true);
        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);

        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        WorksheetPart sheetPart = doc.WorkbookPart!.WorksheetParts.First();
        Assert.NotNull(sheetPart.DrawingsPart);
        Assert.NotEmpty(sheetPart.DrawingsPart!.ChartParts);

        ChartPart chartPart = sheetPart.DrawingsPart.ChartParts.First();
        Assert.NotNull(chartPart.ChartSpace);
        Assert.NotNull(chartPart.ChartSpace.GetFirstChild<Chart>());

        // Drawing references chart relationship
        Assert.Contains(sheetPart.DrawingsPart.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>(), _ => true);
        Assert.Contains(
            sheetPart.DrawingsPart.WorksheetDrawing.Descendants<ChartReference>(),
            cr => !string.IsNullOrEmpty(cr.Id));
    }

    [Fact]
    public void MultiSeries_Formulas_PointAtWorksheetCells()
    {
        using MemoryStream ms = CreateWorkbookWithChart(ExcelChartType.ColumnClustered, multiSeries: true);
        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        ChartPart chartPart = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First();

        List<string> formulas = chartPart.ChartSpace!
            .Descendants<Formula>()
            .Select(f => f.Text)
            .Where(t => !string.IsNullOrEmpty(t))
            .Select(t => t!)
            .ToList();

        Assert.Contains(formulas, f => f.Contains("Sheet1!$A$2:$A$4") || f.Contains("Sheet1!$A$2:$A$4".Replace("$", "")));
        // Categories / values absolute refs
        Assert.Contains(formulas, f => f.IndexOf("$A$", StringComparison.Ordinal) >= 0);
        Assert.Contains(formulas, f => f.IndexOf("$B$", StringComparison.Ordinal) >= 0);
        Assert.Contains(formulas, f => f.IndexOf("$C$", StringComparison.Ordinal) >= 0);

        // Caches present
        Assert.NotEmpty(chartPart.ChartSpace.Descendants<NumberingCache>());
        Assert.NotEmpty(chartPart.ChartSpace.Descendants<StringCache>());
    }

    [Fact]
    public void MultipleCharts_OnSameSheet()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = wb.Worksheets[0];
            sheet["A1"].SetValue("Cat");
            sheet["B1"].SetValue("V1");
            sheet["A2"].SetValue("X");
            sheet["B2"].SetValue(1);
            sheet["A3"].SetValue("Y");
            sheet["B3"].SetValue(2);

            ExcelChart col = sheet.Charts.Add(ExcelChartType.ColumnClustered, "Columns");
            col.AddSeries("B2:B3", "A2:A3", name: "V1");
            col.SetPosition(2, 4, 12, 10);

            ExcelChart line = sheet.Charts.Add(ExcelChartType.Line, "Lines");
            line.AddSeries("B2:B3", "A2:A3", name: "V1");
            line.SetPosition(2, 11, 12, 16);

            Assert.Equal(2, sheet.Charts.Count);
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);

        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        DrawingsPart drawings = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!;
        Assert.Equal(2, drawings.ChartParts.Count());
        Assert.Equal(2, drawings.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>().Count());

        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        Assert.Equal(2, loaded.Worksheets[0].Charts.Count);
        Assert.Contains(loaded.Worksheets[0].Charts, c => c.Type == ExcelChartType.ColumnClustered);
        Assert.Contains(loaded.Worksheets[0].Charts, c => c.Type == ExcelChartType.Line);
        Assert.Equal("Columns", loaded.Worksheets[0].Charts.First(c => c.Type == ExcelChartType.ColumnClustered).Title);
    }

    [Fact]
    public void Chart_Title_Position_Size_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = wb.Worksheets[0];
            sheet["A1"].SetValue("A");
            sheet["B1"].SetValue(10);
            sheet["A2"].SetValue("B");
            sheet["B2"].SetValue(20);

            ExcelChart chart = sheet.Charts.Add(ExcelChartType.BarClustered, "Sales");
            chart.AddSeries("B1:B2", "A1:A2");
            chart.SetPosition(5, 3, 5, 3).SetSize(8, 12);

            Assert.Equal(5, chart.FromRow);
            Assert.Equal(3, chart.FromColumn);
            Assert.Equal(16, chart.ToRow);
            Assert.Equal(10, chart.ToColumn);
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);
        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        ExcelChart chart2 = Assert.Single(loaded.Worksheets[0].Charts);
        Assert.Equal(ExcelChartType.BarClustered, chart2.Type);
        Assert.Equal("Sales", chart2.Title);
        Assert.Equal(5, chart2.FromRow);
        Assert.Equal(3, chart2.FromColumn);
        Assert.Equal(16, chart2.ToRow);
        Assert.Equal(10, chart2.ToColumn);
        Assert.NotEmpty(chart2.Series);
        Assert.Equal("B1:B2", chart2.Series[0].ValuesRange);
        Assert.Equal("A1:A2", chart2.Series[0].CategoriesRange);
    }

    private static MemoryStream CreateWorkbookWithChart(ExcelChartType type, bool multiSeries)
    {
        var ms = new MemoryStream();
        using ExcelWorkbook wb = ExcelWorkbook.Create();
        ExcelWorksheet sheet = wb.Worksheets[0];
        sheet["A1"].SetValue("Month");
        sheet["B1"].SetValue("Actual");
        sheet["C1"].SetValue("Plan");
        sheet["A2"].SetValue("Jan");
        sheet["B2"].SetValue(10);
        sheet["C2"].SetValue(8);
        sheet["A3"].SetValue("Feb");
        sheet["B3"].SetValue(15);
        sheet["C3"].SetValue(12);
        sheet["A4"].SetValue("Mar");
        sheet["B4"].SetValue(7);
        sheet["C4"].SetValue(9);

        ExcelChart chart = sheet.Charts.Add(type, type + " Chart");
        chart.SetPosition(2, 5, 18, 12);
        if (type == ExcelChartType.Scatter)
        {
            chart.AddSeries("B2:B4", "C2:C4", name: "Actual vs Plan");
            if (multiSeries)
            {
                chart.AddSeries("C2:C4", "B2:B4", name: "Plan vs Actual");
            }
        }
        else
        {
            chart.AddSeries("B2:B4", "A2:A4", name: "Actual", nameRange: "B1");
            if (multiSeries)
            {
                chart.AddSeries("C2:C4", "A2:A4", name: "Plan", nameRange: "C1");
            }
        }

        wb.Save(ms);
        ms.Position = 0;
        return ms;
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
