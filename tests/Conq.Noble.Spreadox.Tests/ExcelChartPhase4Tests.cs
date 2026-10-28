using System.Text;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Conq.Noble.Spreadox;

namespace Conq.Noble.Spreadox.Tests;

public class ExcelChartPhase4Tests
{
    [Theory]
    [InlineData(ExcelChartType.Pie)]
    [InlineData(ExcelChartType.Doughnut)]
    [InlineData(ExcelChartType.Area)]
    [InlineData(ExcelChartType.AreaStacked)]
    [InlineData(ExcelChartType.Bubble)]
    [InlineData(ExcelChartType.Radar)]
    [InlineData(ExcelChartType.ColumnStacked)]
    [InlineData(ExcelChartType.BarStacked100)]
    public void AdvancedChartTypes_Write_Validate_And_HaveChartPart(ExcelChartType type)
    {
        using MemoryStream ms = CreateSample(type);
        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);

        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        Assert.NotEmpty(doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts);

        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        Assert.Equal(type, Assert.Single(loaded.Worksheets[0].Charts).Type);
    }

    [Fact]
    public void ComboChart_WithSecondaryAxis_WritesMultiplePlotElements()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = SeedCategories(wb);
            ExcelChart chart = sheet.Charts.Add(ExcelChartType.Combo, "Combo");
            chart.ValueAxis.Title = "Primary";
            chart.SecondaryValueAxis.Title = "Secondary";
            chart.ValueAxis.Minimum = 0;
            chart.ValueAxis.Maximum = 50;
            chart.LegendPosition = ExcelLegendPosition.Bottom;

            chart.AddSeries("B2:B4", "A2:A4", name: "Columns")
                .AsChartType(ExcelChartType.ColumnClustered);

            chart.AddSeries("C2:C4", "A2:A4", name: "Line")
                .AsChartType(ExcelChartType.Line)
                .OnSecondaryAxis();

            chart.DataLabels.Show = true;
            chart.DataLabels.ShowValue = true;
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);

        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        ChartSpace space = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First().ChartSpace!;
        PlotArea plot = space.GetFirstChild<Chart>()!.GetFirstChild<PlotArea>()!;
        Assert.NotNull(plot.GetFirstChild<BarChart>());
        Assert.NotNull(plot.GetFirstChild<LineChart>());
        Assert.Equal(3, plot.Elements<ValueAxis>().Count() + plot.Elements<CategoryAxis>().Count());

        // Secondary axis id 3 referenced
        Assert.Contains(plot.Descendants<AxisId>(), a => a.Val?.Value == 3);

        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        ExcelChart c = Assert.Single(loaded.Worksheets[0].Charts);
        Assert.Equal(ExcelChartType.Combo, c.Type);
        Assert.Equal(2, c.Series.Count);
        Assert.Contains(c.Series, s => s.UseSecondaryAxis);
        Assert.Equal(ExcelLegendPosition.Bottom, c.LegendPosition);
    }

    [Fact]
    public void Pie_DataLabels_And_DoughnutHole()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = SeedCategories(wb);
            ExcelChart pie = sheet.Charts.Add(ExcelChartType.Pie, "Share");
            pie.DataLabels.Show = true;
            pie.DataLabels.ShowPercent = true;
            pie.DataLabels.ShowCategoryName = true;
            pie.AddSeries("B2:B4", "A2:A4", name: "Actual");

            ExcelChart donut = sheet.Charts.Add(ExcelChartType.Doughnut, "Hole");
            donut.DoughnutHoleSize = 60;
            donut.SetPosition(2, 12, 16, 18);
            donut.AddSeries("C2:C4", "A2:A4");
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);
        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        var chartParts = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.ToList();
        Assert.Equal(2, chartParts.Count);
        Assert.Contains(chartParts, p => p.ChartSpace!.Descendants<PieChart>().Any());
        Assert.Contains(chartParts, p => p.ChartSpace!.Descendants<DoughnutChart>().Any());
        DoughnutChart donutNode = chartParts.SelectMany(p => p.ChartSpace!.Descendants<DoughnutChart>()).First();
        Assert.Equal(60, donutNode.GetFirstChild<HoleSize>()!.Val!.Value);

        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        Assert.Equal(60, loaded.Worksheets[0].Charts.First(c => c.Type == ExcelChartType.Doughnut).DoughnutHoleSize);
    }

    [Fact]
    public void Markers_And_AxisOptions_OnLineChart()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = SeedCategories(wb);
            ExcelChart chart = sheet.Charts.Add(ExcelChartType.Line, "Trend");
            chart.CategoryAxis.Title = "Month";
            chart.ValueAxis.Title = "Amount";
            chart.ValueAxis.Minimum = 0;
            chart.ValueAxis.Maximum = 100;
            chart.ValueAxis.ShowMinorGridlines = true;
            ExcelChartSeries s = chart.AddSeries("B2:B4", "A2:A4", name: "Actual");
            s.ShowMarkers = true;
            s.MarkerStyle = ExcelMarkerStyle.Diamond;
            s.DataLabels.Show = true;
            s.DataLabels.ShowValue = true;
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);
        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        ChartSpace space = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First().ChartSpace!;
        Assert.Contains(space.Descendants<Symbol>(), sym => sym.Val?.Value == MarkerStyleValues.Diamond);
        Assert.Contains(space.Descendants<MinAxisValue>(), m => Math.Abs(m.Val!.Value - 0) < 0.001);
        Assert.Contains(space.Descendants<MaxAxisValue>(), m => Math.Abs(m.Val!.Value - 100) < 0.001);
        Assert.NotEmpty(space.Descendants<MinorGridlines>());
        Assert.NotEmpty(space.Descendants<DataLabels>());
    }

    [Fact]
    public void Bubble_RequiresSizeRange()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = SeedCategories(wb);
            ExcelChart chart = sheet.Charts.Add(ExcelChartType.Bubble, "Bubbles");
            chart.AddSeries("B2:B4", "C2:C4", name: "S1").SetBubbleSize("B2:B4");
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);
        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        Assert.NotEmpty(doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First()
            .ChartSpace!.Descendants<BubbleChart>());
    }

    private static MemoryStream CreateSample(ExcelChartType type)
    {
        var ms = new MemoryStream();
        using ExcelWorkbook wb = ExcelWorkbook.Create();
        ExcelWorksheet sheet = SeedCategories(wb);
        ExcelChart chart = sheet.Charts.Add(type, type.ToString());
        if (type == ExcelChartType.Bubble)
        {
            chart.AddSeries("B2:B4", "C2:C4", name: "S1").SetBubbleSize("B2:B4");
        }
        else if (type == ExcelChartType.Scatter)
        {
            chart.AddSeries("B2:B4", "C2:C4", name: "S1");
        }
        else
        {
            chart.AddSeries("B2:B4", "A2:A4", name: "Actual");
            if (type == ExcelChartType.ColumnStacked || type == ExcelChartType.BarStacked100 || type == ExcelChartType.AreaStacked)
            {
                chart.AddSeries("C2:C4", "A2:A4", name: "Plan");
            }
        }

        wb.Save(ms);
        ms.Position = 0;
        return ms;
    }

    private static ExcelWorksheet SeedCategories(ExcelWorkbook wb)
    {
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
        return sheet;
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
