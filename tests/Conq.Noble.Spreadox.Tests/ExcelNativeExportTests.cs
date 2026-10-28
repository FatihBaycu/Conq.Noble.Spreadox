using System.Text;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Conq.Noble.Spreadox;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Conq.Noble.Spreadox.Tests;

public class ExcelNativeExportTests
{
    // 1x1 PNG (red)
    private static readonly byte[] TinyPng =
    {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
        0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00,
        0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
        0x00, 0x00, 0x03, 0x00, 0x01, 0x18, 0xDD, 0x8D, 0xB4, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82
    };

    [Fact]
    public void Picture_RangeAnchor_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = wb.Worksheets[0];
            sheet["A1"].SetValue("Logo");
            string dataUri = "data:image/png;base64," + Convert.ToBase64String(TinyPng);
            ExcelPicture pic = sheet.Pictures.AddFromBase64(dataUri, "A1:B2", lockAspectRatio: true);
            pic.Name = "CompanyLogo";
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);

        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        DrawingsPart drawings = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!;
        Assert.NotEmpty(drawings.ImageParts);
        Xdr.Picture picture = drawings.WorksheetDrawing!.Descendants<Xdr.Picture>().First();
        Assert.True(picture.Descendants<A.PictureLocks>().First().NoChangeAspect?.Value);

        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        ExcelPicture round = Assert.Single(loaded.Worksheets[0].Pictures);
        Assert.Equal(ExcelPictureFormat.Png, round.Format);
        Assert.Equal("A1:B2", round.AnchorRange.A1);
        Assert.True(round.LockAspectRatio);
        Assert.Equal(TinyPng.Length, round.ImageBytes.Length);
    }

    [Fact]
    public void ColorScale_Percentile_WritesAndOpens()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = wb.Worksheets[0];
            sheet["A1"].SetValue(10);
            sheet["A2"].SetValue(50);
            sheet["A3"].SetValue(90);
            sheet.AddColorScale(
                "A1:A3",
                ExcelColor.FromHex("F85149"),
                ExcelColor.FromHex("FFD166"),
                ExcelColor.FromHex("3FB950"));
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);

        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        ConditionalFormattingRule rule = doc.WorkbookPart!.WorksheetParts.First().Worksheet!
            .Descendants<ConditionalFormattingRule>().First();
        Assert.Equal(ConditionalFormatValues.ColorScale, rule.Type?.Value);
        ColorScale scale = rule.GetFirstChild<ColorScale>()!;
        Assert.Equal(3, scale.Elements<ConditionalFormatValueObject>().Count());
        Assert.Contains(scale.Elements<ConditionalFormatValueObject>(), v => v.Val?.Value == "50");

        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        ExcelConditionalFormat cf = Assert.Single(loaded.Worksheets[0].ConditionalFormats);
        Assert.Equal(ExcelConditionalFormatKind.ColorScale, cf.Kind);
        Assert.NotNull(cf.ColorScale);
    }

    [Fact]
    public void SeriesColor_LineOutline_ColumnFill()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = Seed(wb);
            ExcelChart columns = sheet.Charts.Add(ExcelChartType.ColumnClustered, "Cols");
            columns.AddSeries("B2:B4", "A2:A4", name: "Actual").SetColor("2C3E50");
            columns.SetPosition(2, 6, 16, 12);

            ExcelChart line = sheet.Charts.Add(ExcelChartType.LineMarker, "Trend");
            line.AddSeries("C2:C4", "A2:A4", name: "Plan").SetColor("E67E22");
            line.SetPosition(18, 6, 32, 12);
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);

        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        List<ChartPart> parts = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.ToList();
        Assert.Equal(2, parts.Count);
        Assert.Contains(parts, p => p.ChartSpace!.Descendants<A.SolidFill>().Any());
        Assert.Contains(parts, p => p.ChartSpace!.Descendants<A.Outline>().Any());
        Assert.Contains(parts, p => p.ChartSpace!.Descendants<LineChart>().Any());
    }

    [Fact]
    public void DataLabels_Separator_And_LiteralPercentFormat()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = Seed(wb);
            ExcelChart pie = sheet.Charts.Add(ExcelChartType.Pie, "Share");
            pie.DataLabels.Show = true;
            pie.DataLabels.ShowValue = true;
            pie.DataLabels.ShowCategoryName = true;
            pie.DataLabels.Separator = ": ";
            pie.DataLabels.NumberFormat = "0.0\"%\"";
            pie.AddSeries("B2:B4", "A2:A4", name: "Actual");
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);

        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        DataLabels labels = doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ChartParts.First()
            .ChartSpace!.Descendants<DataLabels>().First();
        Assert.Equal(": ", labels.GetFirstChild<Separator>()!.Text);
        C.NumberingFormat fmt = labels.GetFirstChild<C.NumberingFormat>()!;
        Assert.Equal("0.0\"%\"", fmt.FormatCode?.Value);
        Assert.False(fmt.SourceLinked?.Value);
    }

    [Fact]
    public void BoxAndWhisker_WritesChartEx()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = wb.Worksheets[0];
            sheet["A1"].SetValue("G1");
            sheet["B1"].SetValue("G2");
            sheet["C1"].SetValue("G3");
            var rng = new Random(7);
            for (int r = 2; r <= 12; r++)
            {
                sheet[r, 1].SetValue(10 + rng.NextDouble() * 5);
                sheet[r, 2].SetValue(12 + rng.NextDouble() * 6);
                sheet[r, 3].SetValue(8 + rng.NextDouble() * 4);
            }

            ExcelChart chart = sheet.Charts.Add(ExcelChartType.BoxAndWhisker, "BoxPlot");
            chart.ValueAxis.Title = "Value";
            chart.BoxAndWhisker.ShowOutlierPoints = true;
            chart.BoxAndWhisker.ShowMeanMarkers = true;
            chart.BoxAndWhisker.ShowMeanLine = false;
            chart.BoxAndWhisker.ShowInnerPoints = true;
            chart.AddSeries("A1:C12", name: "Groups");
            chart.SetPosition(2, 5, 18, 12);
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);

        ms.Position = 0;
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false);
        ExtendedChartPart ex = Assert.Single(doc.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ExtendedChartParts);
        Assert.NotNull(ex.ChartSpace);
        Cx.Series series = ex.ChartSpace!.Descendants<Cx.Series>().First();
        Assert.Equal(Cx.SeriesLayout.BoxWhisker, series.LayoutId?.Value);
        Cx.SeriesElementVisibilities vis = series.Descendants<Cx.SeriesElementVisibilities>().First();
        Assert.True(vis.Outliers?.Value);
        Assert.True(vis.MeanMarker?.Value);
        Assert.False(vis.MeanLine?.Value);
        Assert.True(vis.Nonoutliers?.Value);

        ms.Position = 0;
        using ExcelWorkbook loaded = ExcelWorkbook.Open(ms);
        ExcelChart loadedChart = Assert.Single(loaded.Worksheets[0].Charts);
        Assert.Equal(ExcelChartType.BoxAndWhisker, loadedChart.Type);
        Assert.True(loadedChart.BoxAndWhisker.ShowOutlierPoints);
    }

    [Fact]
    public void WorksheetName_RejectsColon()
    {
        using ExcelWorkbook wb = ExcelWorkbook.Create();
        Assert.Throws<ArgumentException>(() => wb.AddWorksheet("Bad:Name"));
    }

    [Fact]
    public void Combo_StillValid_WithSeriesColorAndSecondaryAxis()
    {
        using var ms = new MemoryStream();
        using (ExcelWorkbook wb = ExcelWorkbook.Create())
        {
            ExcelWorksheet sheet = Seed(wb);
            ExcelChart combo = sheet.Charts.Add(ExcelChartType.Combo, "Pareto");
            combo.AddSeries("B2:B4", "A2:A4", name: "Count")
                .AsChartType(ExcelChartType.ColumnClustered)
                .SetColor("3498DB");
            combo.AddSeries("C2:C4", "A2:A4", name: "Cum%")
                .AsChartType(ExcelChartType.LineMarker)
                .OnSecondaryAxis()
                .SetColor("E74C3C");
            combo.DataLabels.Show = true;
            combo.DataLabels.ShowValue = true;
            combo.DataLabels.NumberFormat = "0.0\"%\"";
            wb.Save(ms);
        }

        ms.Position = 0;
        AssertNoOpenXmlErrors(ms);
    }

    private static ExcelWorksheet Seed(ExcelWorkbook wb)
    {
        ExcelWorksheet sheet = wb.Worksheets[0];
        sheet["A1"].SetValue("Cat");
        sheet["B1"].SetValue("Actual");
        sheet["C1"].SetValue("Plan");
        sheet["A2"].SetValue("A");
        sheet["B2"].SetValue(10);
        sheet["C2"].SetValue(20);
        sheet["A3"].SetValue("B");
        sheet["B3"].SetValue(30);
        sheet["C3"].SetValue(40);
        sheet["A4"].SetValue("C");
        sheet["B4"].SetValue(50);
        sheet["C4"].SetValue(60);
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
