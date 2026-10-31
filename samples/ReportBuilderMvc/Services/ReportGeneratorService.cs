using System.Globalization;
using Conq.Noble.Spreadox;
using ReportBuilderMvc.Models;

namespace ReportBuilderMvc.Services;

/// <summary>
/// Generates random data and produces Excel/PDF reports using Spreadox.
/// </summary>
public sealed class ReportGeneratorService
{
    private static readonly string[] SampleTitles =
    {
        "Revenue", "Expenses", "Profit", "Users", "Sessions",
        "Conversion", "Bounce Rate", "Orders", "Returns", "Inventory",
        "Temperature", "Humidity", "Pressure", "Velocity", "Output",
        "Quality Score", "Efficiency", "Throughput", "Latency", "Uptime"
    };

    private static readonly string[] DefaultCategories =
    {
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    };

    private static readonly string[] ChartColors =
    {
        "#005293", "#ED7D31", "#5B9BD5", "#70AD47", "#FFC000",
        "#A5A5A5", "#9E480E", "#4472C4", "#264478", "#636363",
        "#255E91", "#9B2335", "#BF8F00", "#44546A", "#C55A11",
        "#2E75B6", "#A9D18E", "#FF6384", "#36A2EB", "#FFCE56"
    };

    /// <summary>
    /// Generate random chart data items.
    /// </summary>
    public List<ChartDataItem> GenerateRandomData(int count)
    {
        var rng = new Random();
        var items = new List<ChartDataItem>();
        var chartTypes = ReportConfigViewModel.AvailableChartTypes;

        for (int i = 0; i < count; i++)
        {
            int pointCount = rng.Next(6, 13);
            var categories = DefaultCategories.Take(pointCount).ToArray();
            var values = new double[pointCount];

            double baseValue = rng.Next(10, 200);
            for (int j = 0; j < pointCount; j++)
            {
                values[j] = Math.Round(baseValue + rng.NextDouble() * baseValue * 0.5 - baseValue * 0.25, 1);
            }

            items.Add(new ChartDataItem
            {
                Title = i < SampleTitles.Length ? SampleTitles[i] : $"Series {i + 1}",
                ChartType = chartTypes[rng.Next(chartTypes.Count)],
                Categories = string.Join(", ", categories),
                Values = string.Join(", ", values.Select(v => v.ToString("F1", CultureInfo.InvariantCulture))),
                Color = ChartColors[i % ChartColors.Length]
            });
        }

        return items;
    }

    /// <summary>
    /// Generate an Excel workbook from the configuration.
    /// </summary>
    public byte[] GenerateExcel(ReportConfigViewModel config)
    {
        using var wb = ExcelWorkbook.Create();

        if (config.Layout == ChartLayout.SinglePage)
        {
            GenerateExcelSingleSheet(wb, config);
        }
        else
        {
            GenerateExcelSeparateSheets(wb, config);
        }

        using var ms = new MemoryStream();
        wb.Save(ms);
        return ms.ToArray();
    }

    /// <summary>
    /// Generate a PDF report from the configuration.
    /// </summary>
    public byte[] GeneratePdf(ReportConfigViewModel config, byte[]? headerLogoBytes = null, byte[]? footerLogoBytes = null)
    {
        using var report = PdfReport.Create(PdfPageSize.A4Landscape);
        report.Margins = PdfMargins.Millimeters(10, 10, 10, 14);

        // Cover page
        report.Cover.Company = config.Cover.Company ?? "Spreadox MVC Sample";
        report.Cover.Title = config.Cover.Title ?? config.ReportTitle;
        report.Cover.Provider = config.Cover.Provider;
        if (config.Cover.IncludeDate)
        {
            report.Cover.Date = DateTime.Now;
        }

        // Header
        if (!string.IsNullOrWhiteSpace(config.Header.Title) || !string.IsNullOrWhiteSpace(config.Header.Text))
        {
            report.Header.Title = config.Header.Title;
            report.Header.Text = config.Header.Text;
            report.Header.FontSize = config.Header.FontSize;
            report.Header.Bold = config.Header.Bold;

            if (!string.IsNullOrWhiteSpace(config.Header.FontFamily))
            {
                report.Header.FontFamily = config.Header.FontFamily;
            }

            if (!string.IsNullOrWhiteSpace(config.Header.TitleColor))
            {
                try
                {
                    report.Header.TitleColor = ExcelColor.FromHex(config.Header.TitleColor);
                }
                catch
                {
                    // Ignore invalid color
                }
            }

            if (headerLogoBytes is { Length: > 0 })
            {
                report.Header.Logo = PdfImage.FromBytes(headerLogoBytes);
            }
        }

        // Footer
        report.Footer.Text = config.Footer.Text;
        report.Footer.PageNumberFormat = config.Footer.PageNumberFormat ?? "Page {0} / {1}";
        if (!string.IsNullOrWhiteSpace(config.Footer.FontFamily))
        {
            report.Footer.FontFamily = config.Footer.FontFamily;
        }

        if (footerLogoBytes is { Length: > 0 })
        {
            report.Footer.Logo = PdfImage.FromBytes(footerLogoBytes);
        }

        // Chart pages
        foreach (var item in config.Items)
        {
            var (categories, values) = ParseChartData(item);
            if (values.Length == 0) continue;

            var page = report.AddChartPage(item.Title);
            page.Chart.Type = item.ChartType;

            var series = page.Chart.AddSeries(item.Title, categories, values);
            try
            {
                series.SetColor(ExcelColor.FromHex(item.Color));
            }
            catch
            {
                // Ignore invalid color
            }

            // Summary stats
            if (values.Length > 0)
            {
                page.Summary.Add("Count", values.Length.ToString());
                page.Summary.Add("Min", values.Min().ToString("F1", CultureInfo.InvariantCulture));
                page.Summary.Add("Max", values.Max().ToString("F1", CultureInfo.InvariantCulture));
                page.Summary.Add("Mean", values.Average().ToString("F1", CultureInfo.InvariantCulture));
                page.Summary.Add("Sum", values.Sum().ToString("F1", CultureInfo.InvariantCulture));
            }
        }

        return report.ToByteArray();
    }

    private void GenerateExcelSingleSheet(ExcelWorkbook wb, ReportConfigViewModel config)
    {
        var sheet = wb.Worksheets[0];
        sheet.Name = "Report";

        int currentRow = 1;

        for (int idx = 0; idx < config.Items.Count; idx++)
        {
            var item = config.Items[idx];
            var (categories, values) = ParseChartData(item);
            if (values.Length == 0) continue;

            int dataStartRow = currentRow;

            // Header row
            sheet[currentRow, 1].SetValue(item.Title);
            sheet[currentRow, 1].Style.Font.Bold = true;
            sheet[currentRow, 1].Style.Font.Size = 14;
            sheet.Merge($"A{currentRow}:B{currentRow}");
            currentRow++;

            // Column headers
            sheet[currentRow, 1].SetValue("Category");
            sheet[currentRow, 2].SetValue("Value");
            sheet[currentRow, 1].Style.Font.Bold = true;
            sheet[currentRow, 2].Style.Font.Bold = true;
            sheet[currentRow, 1].Style.Fill.Pattern = ExcelFillPattern.Solid;
            sheet[currentRow, 1].Style.Fill.ForegroundColor = ExcelColor.FromRgb(220, 230, 241);
            sheet[currentRow, 2].Style.Fill.Pattern = ExcelFillPattern.Solid;
            sheet[currentRow, 2].Style.Fill.ForegroundColor = ExcelColor.FromRgb(220, 230, 241);
            currentRow++;

            int valuesStartRow = currentRow;
            for (int j = 0; j < values.Length; j++)
            {
                sheet[currentRow, 1].SetValue(j < categories.Length ? categories[j] : (j + 1).ToString());
                sheet[currentRow, 2].SetValue(values[j]);
                sheet[currentRow, 2].Style.NumberFormat = "0.00";
                currentRow++;
            }

            int valuesEndRow = currentRow - 1;

            // Add chart
            var chartType = item.ChartType;
            if (chartType == ExcelChartType.Combo)
            {
                chartType = ExcelChartType.ColumnClustered;
            }

            var chart = sheet.Charts.Add(chartType, item.Title);
            chart.AddSeries($"B{valuesStartRow}:B{valuesEndRow}", $"A{valuesStartRow}:A{valuesEndRow}", name: item.Title);
            chart.LegendPosition = ExcelLegendPosition.Bottom;

            int chartFromCol = 4;
            int chartToCol = 14;
            chart.SetPosition(dataStartRow, chartFromCol, dataStartRow + 14, chartToCol);

            currentRow += 2; // spacing
        }

        sheet.SetColumnWidth("A", 14);
        sheet.SetColumnWidth("B", 12);
    }

    private void GenerateExcelSeparateSheets(ExcelWorkbook wb, ReportConfigViewModel config)
    {
        // Remove default sheet if we'll add named ones
        bool firstSheet = true;

        for (int idx = 0; idx < config.Items.Count; idx++)
        {
            var item = config.Items[idx];
            var (categories, values) = ParseChartData(item);
            if (values.Length == 0) continue;

            ExcelWorksheet sheet;
            if (firstSheet)
            {
                sheet = wb.Worksheets[0];
                sheet.Name = SanitizeSheetName(item.Title);
                firstSheet = false;
            }
            else
            {
                sheet = wb.AddWorksheet(SanitizeSheetName(item.Title, idx));
            }

            // Column headers
            sheet["A1"].SetValue("Category");
            sheet["B1"].SetValue("Value");
            sheet["A1"].Style.Font.Bold = true;
            sheet["B1"].Style.Font.Bold = true;
            sheet["A1"].Style.Fill.Pattern = ExcelFillPattern.Solid;
            sheet["A1"].Style.Fill.ForegroundColor = ExcelColor.FromRgb(220, 230, 241);
            sheet["B1"].Style.Fill.Pattern = ExcelFillPattern.Solid;
            sheet["B1"].Style.Fill.ForegroundColor = ExcelColor.FromRgb(220, 230, 241);

            for (int j = 0; j < values.Length; j++)
            {
                int row = j + 2;
                sheet[row, 1].SetValue(j < categories.Length ? categories[j] : (j + 1).ToString());
                sheet[row, 2].SetValue(values[j]);
                sheet[row, 2].Style.NumberFormat = "0.00";
            }

            int endRow = values.Length + 1;

            // Add chart
            var chartType = item.ChartType;
            if (chartType == ExcelChartType.Combo)
            {
                chartType = ExcelChartType.ColumnClustered;
            }

            var chart = sheet.Charts.Add(chartType, item.Title);
            chart.AddSeries($"B2:B{endRow}", $"A2:A{endRow}", name: item.Title);
            chart.LegendPosition = ExcelLegendPosition.Bottom;
            chart.SetPosition(2, 4, 18, 14);

            // Summary stats below chart
            int summaryRow = endRow + 2;
            sheet[summaryRow, 1].SetValue("Statistics");
            sheet[summaryRow, 1].Style.Font.Bold = true;
            sheet[summaryRow + 1, 1].SetValue("Count");
            sheet[summaryRow + 1, 2].SetValue(values.Length);
            sheet[summaryRow + 2, 1].SetValue("Min");
            sheet[summaryRow + 2, 2].SetValue(values.Min());
            sheet[summaryRow + 3, 1].SetValue("Max");
            sheet[summaryRow + 3, 2].SetValue(values.Max());
            sheet[summaryRow + 4, 1].SetValue("Mean");
            sheet[summaryRow + 4, 2].SetValue(Math.Round(values.Average(), 2));

            sheet.SetColumnWidth("A", 14);
            sheet.SetColumnWidth("B", 12);
            sheet.FreezePanes(1);
        }
    }

    private static (string[] categories, double[] values) ParseChartData(ChartDataItem item)
    {
        var categories = (item.Categories ?? "")
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToArray();

        var values = (item.Values ?? "")
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(v => double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0.0)
            .ToArray();

        return (categories, values);
    }

    private static string SanitizeSheetName(string name, int index = 0)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return $"Sheet{index + 1}";
        }

        // Excel sheet name max 31 chars, no special chars
        string clean = new string(name.Where(c => c != '\\' && c != '/' && c != '?' && c != '*'
            && c != '[' && c != ']' && c != ':').ToArray());

        if (clean.Length > 28)
        {
            clean = clean[..28];
        }

        return string.IsNullOrWhiteSpace(clean) ? $"Sheet{index + 1}" : clean;
    }
}

