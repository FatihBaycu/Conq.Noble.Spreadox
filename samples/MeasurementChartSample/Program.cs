using Conq.Noble.Spreadox;

// Classic measurement-report scenario: stations × values with target line (combo chart).
string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SpreadoxSamples");
Directory.CreateDirectory(outputDir);
string reportPath = Path.Combine(outputDir, "measurement-report.xlsx");
string streamPath = Path.Combine(outputDir, "measurement-stream-sample.xlsx");

WriteMeasurementReport(reportPath);
WriteStreamingExcerpt(streamPath);

Console.WriteLine("Wrote:");
Console.WriteLine("  " + reportPath);
Console.WriteLine("  " + streamPath);

static void WriteMeasurementReport(string path)
{
    using ExcelWorkbook wb = ExcelWorkbook.Create();
    ExcelWorksheet sheet = wb.Worksheets[0];
    sheet.Name = "Measurements";

    // Header
    sheet["A1"].SetValue("Station");
    sheet["B1"].SetValue("Actual");
    sheet["C1"].SetValue("Target");
    sheet["D1"].SetValue("Delta");
    foreach (string col in new[] { "A1", "B1", "C1", "D1" })
    {
        sheet[col].Style.Font.Bold = true;
        sheet[col].Style.Fill.Pattern = ExcelFillPattern.Solid;
        sheet[col].Style.Fill.ForegroundColor = ExcelColor.FromRgb(220, 230, 241);
    }

    // Sample measurement points (API-style station readings)
    (string Station, double Actual, double Target)[] rows =
    {
        ("ST-01", 12.4, 12.0),
        ("ST-02", 11.8, 12.0),
        ("ST-03", 13.1, 12.5),
        ("ST-04", 12.0, 12.5),
        ("ST-05", 14.2, 13.0),
        ("ST-06", 13.5, 13.0),
        ("ST-07", 12.7, 12.5),
        ("ST-08", 11.9, 12.0),
    };

    for (int i = 0; i < rows.Length; i++)
    {
        int r = i + 2;
        sheet[r, 1].SetValue(rows[i].Station);
        sheet[r, 2].SetValue(rows[i].Actual);
        sheet[r, 3].SetValue(rows[i].Target);
        sheet.SetFormula($"D{r}", $"B{r}-C{r}", cachedValue: rows[i].Actual - rows[i].Target);
        sheet[r, 2].Style.NumberFormat = "0.00";
        sheet[r, 3].Style.NumberFormat = "0.00";
        sheet[r, 4].Style.NumberFormat = "0.00";
    }

    sheet.SetColumnWidth("A", 12);
    sheet.SetColumnWidth("B", 10);
    sheet.SetColumnWidth("C", 10);
    sheet.SetColumnWidth("D", 10);
    sheet.FreezePanes(1);
    sheet.AutoFilterRange = "A1:D9";
    sheet.AddConditionalFormat("D2:D9", ConditionalOperator.GreaterThan, "0.5", ExcelColor.FromRgb(255, 199, 206));

    // Combo: clustered columns for Actual + line for Target on secondary axis
    ExcelChart chart = sheet.Charts.Add(ExcelChartType.Combo, "Station measurements");
    chart.LegendPosition = ExcelLegendPosition.Bottom;
    chart.ValueAxis.Title = "Actual";
    chart.SecondaryValueAxis.Title = "Target";
    chart.CategoryAxis.Title = "Station";
    chart.DataLabels.Show = false;

    chart.AddSeries("B2:B9", "A2:A9", name: "Actual")
        .AsChartType(ExcelChartType.ColumnClustered);

    chart.AddSeries("C2:C9", "A2:A9", name: "Target")
        .AsChartType(ExcelChartType.Line)
        .OnSecondaryAxis();

    chart.SetPosition(2, 6, 18, 14);

    // Simple pie of absolute actuals for share view
    ExcelChart pie = sheet.Charts.Add(ExcelChartType.Pie, "Actual share");
    pie.DataLabels.Show = true;
    pie.DataLabels.ShowPercent = true;
    pie.DataLabels.ShowCategoryName = true;
    pie.AddSeries("B2:B9", "A2:A9", name: "Actual");
    pie.SetPosition(2, 15, 18, 22);

    wb.AddNamedRange("ActualValues", "'Measurements'!$B$2:$B$9");
    wb.Save(path);
}

static void WriteStreamingExcerpt(string path)
{
    // Streaming path for bulk measurement dumps (no in-memory cell grid / no charts).
    using ExcelStreamingWorkbook sw = ExcelStreamingWorkbook.Create(path);
    ExcelCellStyle header = sw.CreateStyle();
    header.Font.Bold = true;

    ExcelStreamingWorksheet sheet = sw.AddWorksheet("Bulk");
    sheet.SetColumnWidth(1, 14);
    sheet.SetColumnWidth(2, 12);
    sheet.WriteRow(1, row =>
    {
        row.Text(1, "Station", header);
        row.Text(2, "Value", header);
    });

    for (int i = 1; i <= 1000; i++)
    {
        int n = i;
        sheet.WriteRow(n + 1, row =>
        {
            row.Text(1, "ST-" + n.ToString("D4"));
            row.Number(2, 10 + (n % 50) * 0.1);
        });
    }
}
