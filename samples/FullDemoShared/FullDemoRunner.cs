using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using Conq.Noble.Spreadox;

namespace Conq.Noble.Spreadox.Samples
{
    /// <summary>
    /// End-to-end library walkthrough used by the net8 and net48 console samples.
    /// </summary>
    public static class FullDemoRunner
    {
        private static readonly string[] Months =
        {
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        };

        private static readonly string[] Products =
        {
            "Alpha", "Beta", "Gamma", "Delta", "Epsilon", "Zeta", "Eta", "Theta"
        };

        private static readonly string[] Regions = { "North", "South", "East", "West" };

        private static readonly string[] RadarAxes =
        {
            "Speed", "Quality", "Cost", "Delivery", "Support", "Flexibility"
        };

        private static readonly byte[] TinyPng =
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
            0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00,
            0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x03, 0x00, 0x01, 0x18, 0xDD, 0x8D, 0xB4, 0x00, 0x00, 0x00,
            0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82
        };

        private static readonly ExcelColor HeaderFill = ExcelColor.FromRgb(0, 82, 147);
        private static readonly ExcelColor AltRowFill = ExcelColor.FromRgb(242, 246, 250);
        private static readonly ExcelColor AccentFill = ExcelColor.FromRgb(220, 230, 241);
        private static readonly ExcelColor AlertFill = ExcelColor.FromRgb(255, 199, 206);
        private static readonly ExcelColor OkFill = ExcelColor.FromRgb(198, 239, 206);

        public static int Run(string runtimeLabel)
        {
            int seed = Environment.TickCount;
            var rng = new Random(seed);

            string outputDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "SpreadoxSamples");
            Directory.CreateDirectory(outputDir);

            string safeLabel = SanitizeFileToken(runtimeLabel);
            string workbookPath = Path.Combine(outputDir, "full-demo-" + safeLabel + ".xlsx");
            string streamPath = Path.Combine(outputDir, "full-demo-stream-" + safeLabel + ".xlsx");

            Console.WriteLine("Conq.Noble.Spreadox full demo");
            Console.WriteLine("  Runtime : " + runtimeLabel);
            Console.WriteLine("  Seed    : " + seed.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine();

            WriteWorkbook(workbookPath, rng, runtimeLabel, seed);
            WriteStreamingWorkbook(streamPath, rng);

            VerifyRoundTrip(workbookPath);
            VerifyRoundTrip(streamPath);

            Console.WriteLine();
            Console.WriteLine("Wrote:");
            Console.WriteLine("  " + workbookPath);
            Console.WriteLine("  " + streamPath);
            return 0;
        }

        private static void WriteWorkbook(string path, Random rng, string runtimeLabel, int seed)
        {
            using ExcelWorkbook wb = ExcelWorkbook.Create();
            wb.RenameWorksheet("Sheet1", "KitchenSink");
            ExcelWorksheet scratch = wb.AddWorksheet("Scratch");
            scratch["A1"].SetValue("removed");
            wb.RemoveWorksheet(scratch);

            WriteKitchenSink(wb["KitchenSink"], rng, runtimeLabel, seed);
            WriteColumnCharts(wb.AddWorksheet("ColumnCharts"), rng);
            WriteBarCharts(wb.AddWorksheet("BarCharts"), rng);
            WriteLineCharts(wb.AddWorksheet("LineCharts"), rng);
            WriteAreaCharts(wb.AddWorksheet("AreaCharts"), rng);
            WritePieCharts(wb.AddWorksheet("PieCharts"), rng);
            WriteXyCharts(wb.AddWorksheet("XYCharts"), rng);
            WriteRadarChart(wb.AddWorksheet("RadarChart"), rng);
            WriteComboChart(wb.AddWorksheet("ComboChart"), rng);
            WriteBoxPlot(wb.AddWorksheet("BoxPlot"), rng);

            wb.AddNamedRange("MonthlyNorth", "'ColumnCharts'!$B$2:$B$13");
            wb.AddNamedRange("ShareValues", "'PieCharts'!$B$2:$B$9");
            wb.Save(path);
        }

        private static void WriteKitchenSink(ExcelWorksheet sheet, Random rng, string runtimeLabel, int seed)
        {
            sheet.SetRowHeight(1, 28);
            sheet.SetColumnWidth("A", 18);
            sheet.SetColumnWidth("B", 16);
            sheet.SetColumnWidth("C", 14);
            sheet.SetColumnWidth("D", 14);
            sheet.SetColumnWidth("E", 14);
            sheet.SetColumnWidth("F", 22);

            sheet["A1"].SetValue("Conq.Noble.Spreadox kitchen sink — " + runtimeLabel);
            sheet["A1"].Style.Font.Bold = true;
            sheet["A1"].Style.Font.Size = 16;
            sheet["A1"].Style.Font.Name = "Calibri";
            sheet["A1"].Style.Font.Color = ExcelColor.White;
            sheet["A1"].Style.Fill.Pattern = ExcelFillPattern.Solid;
            sheet["A1"].Style.Fill.ForegroundColor = HeaderFill;
            sheet["A1"].Style.Alignment.Horizontal = ExcelHorizontalAlignment.Left;
            sheet["A1"].Style.Alignment.Vertical = ExcelVerticalAlignment.Center;
            sheet.Merge("A1:F1");
            sheet.Pictures.AddFromBase64(
                    "data:image/png;base64," + Convert.ToBase64String(TinyPng),
                    "H1:I2",
                    lockAspectRatio: true)
                .Name = "CompanyLogo";

            sheet["A2"].SetValue("Random seed");
            sheet["B2"].SetValue(seed);
            sheet["C2"].SetValue("Generated");
            sheet["D2"].SetValue(DateTime.Now);
            sheet["D2"].Style.NumberFormat = NumberFormats.DateTime;

            // Cell types
            sheet["A4"].SetValue("Cell types");
            ApplySectionStyle(sheet["A4"]);
            sheet.Merge("A4:H4");

            WriteHeaderRow(sheet, 5, "String", "Integer", "Double", "Decimal", "Boolean", "DateTime", "Empty", "Formula");
            sheet["A6"].SetValue("hello");
            sheet["B6"].SetValue(rng.Next(10, 500));
            sheet["C6"].SetValue(Math.Round(rng.NextDouble() * 1000, 4));
            sheet["D6"].SetValue((decimal)Math.Round(rng.NextDouble() * 250, 2));
            sheet["E6"].SetValue(rng.Next(0, 2) == 1);
            sheet["F6"].SetValue(DateTime.Today.AddDays(-rng.Next(0, 40)).AddHours(rng.Next(8, 18)));
            sheet["F6"].Style.NumberFormat = NumberFormats.DateTime;
            sheet["G6"].Clear();
            sheet.SetValue("H6", null);
            sheet.SetFormula("H6", "B6+C6", cachedValue: Convert.ToDouble(sheet["B6"].Value) + Convert.ToDouble(sheet["C6"].Value));
            sheet["C6"].Style.NumberFormat = NumberFormats.ThousandsTwoDecimals;
            sheet["D6"].Style.NumberFormat = NumberFormats.Currency;
            sheet["B6"].Style.NumberFormat = NumberFormats.Integer;

            // Formats / percent / borders
            sheet["A8"].SetValue("Formats");
            ApplySectionStyle(sheet["A8"]);
            sheet.Merge("A8:C8");
            sheet["A9"].SetValue("Percent");
            sheet["B9"].SetValue(Math.Round(rng.NextDouble(), 4));
            sheet["B9"].Style.NumberFormat = NumberFormats.PercentTwoDecimals;
            sheet["A10"].SetValue("Thousands");
            sheet["B10"].SetValue(rng.Next(1500, 25000) + rng.NextDouble());
            sheet["B10"].Style.NumberFormat = NumberFormats.ThousandsTwoDecimals;
            sheet["B9"].Style.Borders.SetAll(ExcelBorderStyle.Thin, ExcelColor.Gray);
            sheet["B10"].Style.Font.Italic = true;

            // Bulk object insert
            sheet["A12"].SetValue("InsertObjects — station readings");
            ApplySectionStyle(sheet["A12"]);
            sheet.Merge("A12:E12");

            var readings = new List<StationReading>();
            for (int i = 0; i < 12; i++)
            {
                double target = Math.Round(10 + rng.NextDouble() * 4, 2);
                double actual = Math.Round(target + (rng.NextDouble() - 0.45) * 2.4, 2);
                readings.Add(new StationReading
                {
                    Station = "ST-" + (i + 1).ToString("D2", CultureInfo.InvariantCulture),
                    Actual = actual,
                    Target = target,
                    InSpec = Math.Abs(actual - target) < 0.8,
                    TakenAt = DateTime.Today.AddDays(-11 + i)
                });
            }

            sheet.InsertObjects(readings, startRow: 13, startColumn: 1, includeHeaders: true);
            StyleHeaderRow(sheet, 13, 5);
            int lastReading = 13 + readings.Count;
            for (int r = 14; r <= lastReading; r++)
            {
                sheet[r, 2].Style.NumberFormat = NumberFormats.TwoDecimals;
                sheet[r, 3].Style.NumberFormat = NumberFormats.TwoDecimals;
                sheet[r, 5].Style.NumberFormat = NumberFormats.Date;
                if ((r % 2) == 0)
                {
                    ApplyAltRow(sheet, r, 5);
                }
            }

            sheet["F13"].SetValue("Delta");
            StyleHeaderCell(sheet["F13"]);
            for (int r = 14; r <= lastReading; r++)
            {
                sheet.SetFormula("F" + r.ToString(CultureInfo.InvariantCulture),
                    "B" + r.ToString(CultureInfo.InvariantCulture) + "-C" + r.ToString(CultureInfo.InvariantCulture),
                    cachedValue: Convert.ToDouble(sheet[r, 2].Value) - Convert.ToDouble(sheet[r, 3].Value));
                sheet[r, 6].Style.NumberFormat = NumberFormats.TwoDecimals;
            }

            sheet.AddConditionalFormat("F14:F" + lastReading.ToString(CultureInfo.InvariantCulture),
                ConditionalOperator.GreaterThan, "0.5", AlertFill);
            sheet.AddConditionalFormat("F14:F" + lastReading.ToString(CultureInfo.InvariantCulture),
                ConditionalOperator.LessThan, "-0.5", ExcelColor.FromRgb(189, 215, 238));
            sheet.AddColorScale(
                "B14:B" + lastReading.ToString(CultureInfo.InvariantCulture),
                ExcelColor.FromHex("F85149"),
                ExcelColor.FromHex("FFD166"),
                ExcelColor.FromHex("3FB950"));
            sheet.AutoFilterRange = "A13:F" + lastReading.ToString(CultureInfo.InvariantCulture);
            sheet.FreezePanes(13);

            // DataTable insert
            int tableStart = lastReading + 3;
            sheet[tableStart, 1].SetValue("InsertDataTable — random orders");
            ApplySectionStyle(sheet[tableStart, 1]);
            sheet.Merge("A" + tableStart.ToString(CultureInfo.InvariantCulture) + ":E" + tableStart.ToString(CultureInfo.InvariantCulture));

            DataTable orders = BuildOrdersTable(rng);
            int dataStart = tableStart + 1;
            sheet.InsertDataTable(orders, startRow: dataStart, startColumn: 1, includeHeaders: true);
            StyleHeaderRow(sheet, dataStart, orders.Columns.Count);
            int tableLast = dataStart + orders.Rows.Count;
            for (int r = dataStart + 1; r <= tableLast; r++)
            {
                sheet[r, 3].Style.NumberFormat = NumberFormats.Integer;
                sheet[r, 4].Style.NumberFormat = NumberFormats.Currency;
                sheet[r, 5].Style.NumberFormat = NumberFormats.Date;
            }

            // Jagged Insert
            int insertStart = tableLast + 3;
            sheet[insertStart, 1].SetValue("Insert — object arrays");
            ApplySectionStyle(sheet[insertStart, 1]);
            sheet.Merge("A" + insertStart.ToString(CultureInfo.InvariantCulture) + ":C" + insertStart.ToString(CultureInfo.InvariantCulture));
            sheet.Insert(
                new[]
                {
                    new object?[] { "Region", "Units", "Share" },
                    new object?[] { "North", rng.Next(80, 140), 0.31 },
                    new object?[] { "South", rng.Next(70, 130), 0.24 },
                    new object?[] { "East", rng.Next(60, 120), 0.22 },
                    new object?[] { "West", rng.Next(50, 110), 0.23 },
                },
                startRow: insertStart + 1,
                startColumn: 1);
            StyleHeaderRow(sheet, insertStart + 1, 3);
            for (int r = insertStart + 2; r <= insertStart + 5; r++)
            {
                sheet[r, 3].Style.NumberFormat = NumberFormats.PercentTwoDecimals;
            }
        }

        private static void WriteColumnCharts(ExcelWorksheet sheet, Random rng)
        {
            WriteMonthlyRegionTable(sheet, rng, "Monthly volume by region (random)");
            AddStandardChart(sheet, ExcelChartType.ColumnClustered, "Column clustered", 2, 7, 16, 14);
            AddStackedChart(sheet, ExcelChartType.ColumnStacked, "Column stacked", 18, 7, 32, 14);
            AddStackedChart(sheet, ExcelChartType.ColumnStacked100, "Column stacked 100%", 34, 7, 48, 14);
        }

        private static void WriteBarCharts(ExcelWorksheet sheet, Random rng)
        {
            WriteMonthlyRegionTable(sheet, rng, "Monthly volume by region (random)");
            AddStandardChart(sheet, ExcelChartType.BarClustered, "Bar clustered", 2, 7, 16, 14);
            AddStackedChart(sheet, ExcelChartType.BarStacked, "Bar stacked", 18, 7, 32, 14);
            AddStackedChart(sheet, ExcelChartType.BarStacked100, "Bar stacked 100%", 34, 7, 48, 14);
        }

        private static void WriteLineCharts(ExcelWorksheet sheet, Random rng)
        {
            WriteMonthlyRegionTable(sheet, rng, "Monthly trend (random)");
            ExcelChart line = AddStandardChart(sheet, ExcelChartType.Line, "Line", 2, 7, 16, 14);
            line.CategoryAxis.Title = "Month";
            line.ValueAxis.Title = "Units";
            line.ValueAxis.Minimum = 0;
            line.ValueAxis.ShowMinorGridlines = true;
            ExcelChartSeries first = line.Series[0];
            first.ShowMarkers = true;
            first.MarkerStyle = ExcelMarkerStyle.Diamond;
            first.DataLabels.Show = true;
            first.DataLabels.ShowValue = true;

            AddStackedChart(sheet, ExcelChartType.LineStacked, "Line stacked", 18, 7, 32, 14);
            AddStackedChart(sheet, ExcelChartType.LineStacked100, "Line stacked 100%", 34, 7, 48, 14);
        }

        private static void WriteAreaCharts(ExcelWorksheet sheet, Random rng)
        {
            WriteMonthlyRegionTable(sheet, rng, "Monthly area (random)");
            AddStandardChart(sheet, ExcelChartType.Area, "Area", 2, 7, 16, 14);
            AddStackedChart(sheet, ExcelChartType.AreaStacked, "Area stacked", 18, 7, 32, 14);
            AddStackedChart(sheet, ExcelChartType.AreaStacked100, "Area stacked 100%", 34, 7, 48, 14);
        }

        private static void WritePieCharts(ExcelWorksheet sheet, Random rng)
        {
            sheet["A1"].SetValue("Product share (random)");
            ApplySectionStyle(sheet["A1"]);
            sheet.Merge("A1:B1");
            WriteHeaderRow(sheet, 2, "Product", "Revenue");

            for (int i = 0; i < Products.Length; i++)
            {
                int r = i + 3;
                sheet[r, 1].SetValue(Products[i]);
                sheet[r, 2].SetValue(rng.Next(40, 280) * 100);
                sheet[r, 2].Style.NumberFormat = NumberFormats.Thousands;
            }

            sheet.SetColumnWidth("A", 14);
            sheet.SetColumnWidth("B", 12);
            StyleHeaderRow(sheet, 2, 2);

            ExcelChart pie = sheet.Charts.Add(ExcelChartType.Pie, "Pie");
            pie.DataLabels.Show = true;
            pie.DataLabels.ShowValue = true;
            pie.DataLabels.ShowCategoryName = true;
            pie.DataLabels.Separator = ": ";
            pie.DataLabels.NumberFormat = "0.0\"%\"";
            pie.LegendPosition = ExcelLegendPosition.Bottom;
            pie.AddSeries("B3:B10", "A3:A10", name: "Revenue");
            pie.SetPosition(2, 4, 18, 11);

            ExcelChart donut = sheet.Charts.Add(ExcelChartType.Doughnut, "Doughnut");
            donut.DoughnutHoleSize = 60;
            donut.DataLabels.Show = true;
            donut.DataLabels.ShowPercent = true;
            donut.AddSeries("B3:B10", "A3:A10", name: "Revenue");
            donut.SetPosition(20, 4, 36, 11);
        }

        private static void WriteXyCharts(ExcelWorksheet sheet, Random rng)
        {
            sheet["A1"].SetValue("XY / bubble sample (random)");
            ApplySectionStyle(sheet["A1"]);
            sheet.Merge("A1:D1");
            WriteHeaderRow(sheet, 2, "Point", "X", "Y", "Size");

            for (int i = 0; i < 12; i++)
            {
                int r = i + 3;
                sheet[r, 1].SetValue("P" + (i + 1).ToString("D2", CultureInfo.InvariantCulture));
                sheet[r, 2].SetValue(Math.Round(rng.NextDouble() * 100, 1));
                sheet[r, 3].SetValue(Math.Round(20 + rng.NextDouble() * 80, 1));
                sheet[r, 4].SetValue(rng.Next(8, 45));
            }

            sheet.SetColumnWidth("A", 10);
            StyleHeaderRow(sheet, 2, 4);

            ExcelChart scatter = sheet.Charts.Add(ExcelChartType.Scatter, "Scatter");
            scatter.CategoryAxis.Title = "X";
            scatter.ValueAxis.Title = "Y";
            ExcelChartSeries scatterSeries = scatter.AddSeries("C3:C14", "B3:B14", name: "Points");
            scatterSeries.ShowMarkers = true;
            scatterSeries.MarkerStyle = ExcelMarkerStyle.Circle;
            scatter.SetPosition(2, 6, 18, 13);

            ExcelChart bubble = sheet.Charts.Add(ExcelChartType.Bubble, "Bubble");
            bubble.CategoryAxis.Title = "X";
            bubble.ValueAxis.Title = "Y";
            bubble.AddSeries("C3:C14", "B3:B14", name: "Bubbles").SetBubbleSize("D3:D14");
            bubble.SetPosition(20, 6, 36, 13);
        }

        private static void WriteRadarChart(ExcelWorksheet sheet, Random rng)
        {
            sheet["A1"].SetValue("Capability scores (random 1–10)");
            ApplySectionStyle(sheet["A1"]);
            sheet.Merge("A1:C1");
            WriteHeaderRow(sheet, 2, "Axis", "Product A", "Product B");

            for (int i = 0; i < RadarAxes.Length; i++)
            {
                int r = i + 3;
                sheet[r, 1].SetValue(RadarAxes[i]);
                sheet[r, 2].SetValue(rng.Next(4, 11));
                sheet[r, 3].SetValue(rng.Next(4, 11));
            }

            sheet.SetColumnWidth("A", 14);
            StyleHeaderRow(sheet, 2, 3);

            ExcelChart radar = sheet.Charts.Add(ExcelChartType.Radar, "Radar");
            radar.LegendPosition = ExcelLegendPosition.Bottom;
            radar.AddSeries("B3:B8", "A3:A8", nameRange: "B2");
            radar.AddSeries("C3:C8", "A3:A8", nameRange: "C2");
            radar.SetPosition(2, 5, 18, 12);
        }

        private static void WriteComboChart(ExcelWorksheet sheet, Random rng)
        {
            sheet["A1"].SetValue("Actual vs target with yield (random)");
            ApplySectionStyle(sheet["A1"]);
            sheet.Merge("A1:D1");
            WriteHeaderRow(sheet, 2, "Station", "Actual", "Target", "Yield");

            for (int i = 0; i < 8; i++)
            {
                int r = i + 3;
                double target = Math.Round(12 + rng.NextDouble() * 3, 2);
                double actual = Math.Round(target + (rng.NextDouble() - 0.4) * 2, 2);
                sheet[r, 1].SetValue("ST-" + (i + 1).ToString("D2", CultureInfo.InvariantCulture));
                sheet[r, 2].SetValue(actual);
                sheet[r, 3].SetValue(target);
                sheet[r, 4].SetValue(Math.Round(actual / target, 4));
                sheet[r, 2].Style.NumberFormat = NumberFormats.TwoDecimals;
                sheet[r, 3].Style.NumberFormat = NumberFormats.TwoDecimals;
                sheet[r, 4].Style.NumberFormat = NumberFormats.PercentTwoDecimals;
            }

            sheet.SetColumnWidth("A", 12);
            StyleHeaderRow(sheet, 2, 4);
            sheet.AddConditionalFormat("D3:D10", ConditionalOperator.LessThan, "0.95", AlertFill);
            sheet.AddConditionalFormat("D3:D10", ConditionalOperator.GreaterThanOrEqual, "1", OkFill);

            ExcelChart combo = sheet.Charts.Add(ExcelChartType.Combo, "Combo — columns + line");
            combo.LegendPosition = ExcelLegendPosition.Bottom;
            combo.CategoryAxis.Title = "Station";
            combo.ValueAxis.Title = "Measurement";
            combo.SecondaryValueAxis.Title = "Yield";
            combo.ValueAxis.Minimum = 0;
            combo.AddSeries("B3:B10", "A3:A10", name: "Actual")
                .AsChartType(ExcelChartType.ColumnClustered)
                .SetColor("2C3E50");
            combo.AddSeries("C3:C10", "A3:A10", name: "Target")
                .AsChartType(ExcelChartType.LineMarker)
                .SetColor("E67E22");
            combo.AddSeries("D3:D10", "A3:A10", name: "Yield")
                .AsChartType(ExcelChartType.Line)
                .OnSecondaryAxis()
                .SetColor("27AE60");
            combo.Series[1].ShowMarkers = true;
            combo.Series[1].MarkerStyle = ExcelMarkerStyle.Triangle;
            combo.SetPosition(2, 6, 20, 14);
        }

        private static void WriteBoxPlot(ExcelWorksheet sheet, Random rng)
        {
            sheet["A1"].SetValue("Group");
            sheet["B1"].SetValue("North");
            sheet["C1"].SetValue("South");
            sheet["D1"].SetValue("East");
            StyleHeaderRow(sheet, 1, 4);

            for (int r = 2; r <= 16; r++)
            {
                sheet[r, 1].SetValue("N" + (r - 1).ToString(CultureInfo.InvariantCulture));
                sheet[r, 2].SetValue(Math.Round(10 + rng.NextDouble() * 8, 2));
                sheet[r, 3].SetValue(Math.Round(12 + rng.NextDouble() * 6, 2));
                sheet[r, 4].SetValue(Math.Round(8 + rng.NextDouble() * 10, 2));
            }

            ExcelChart box = sheet.Charts.Add(ExcelChartType.BoxAndWhisker, "Box and whisker");
            box.ValueAxis.Title = "Value";
            box.BoxAndWhisker.ShowOutlierPoints = true;
            box.BoxAndWhisker.ShowMeanMarkers = true;
            box.BoxAndWhisker.ShowMeanLine = false;
            box.BoxAndWhisker.ShowInnerPoints = true;
            box.AddSeries("B1:D16", name: "Regions");
            box.SetPosition(2, 6, 20, 14);
        }

        private static void WriteMonthlyRegionTable(ExcelWorksheet sheet, Random rng, string title)
        {
            sheet["A1"].SetValue(title);
            ApplySectionStyle(sheet["A1"]);
            sheet.Merge("A1:E1");
            WriteHeaderRow(sheet, 2, "Month", Regions[0], Regions[1], Regions[2], Regions[3]);

            for (int i = 0; i < Months.Length; i++)
            {
                int r = i + 3;
                sheet[r, 1].SetValue(Months[i]);
                for (int c = 0; c < Regions.Length; c++)
                {
                    sheet[r, c + 2].SetValue(rng.Next(40, 180));
                }

                if ((r % 2) == 0)
                {
                    ApplyAltRow(sheet, r, 5);
                }
            }

            sheet.SetColumnWidth("A", 10);
            StyleHeaderRow(sheet, 2, 5);
            sheet.FreezePanes(2);
        }

        private static ExcelChart AddStandardChart(
            ExcelWorksheet sheet,
            ExcelChartType type,
            string title,
            int fromRow,
            int fromCol,
            int toRow,
            int toCol)
        {
            ExcelChart chart = sheet.Charts.Add(type, title);
            chart.LegendPosition = ExcelLegendPosition.Bottom;
            chart.AddSeries("B3:B14", "A3:A14", nameRange: "B2");
            chart.AddSeries("C3:C14", "A3:A14", nameRange: "C2");
            chart.SetPosition(fromRow, fromCol, toRow, toCol);
            return chart;
        }

        private static ExcelChart AddStackedChart(
            ExcelWorksheet sheet,
            ExcelChartType type,
            string title,
            int fromRow,
            int fromCol,
            int toRow,
            int toCol)
        {
            ExcelChart chart = sheet.Charts.Add(type, title);
            chart.LegendPosition = ExcelLegendPosition.Bottom;
            chart.AddSeries("B3:B14", "A3:A14", nameRange: "B2");
            chart.AddSeries("C3:C14", "A3:A14", nameRange: "C2");
            chart.AddSeries("D3:D14", "A3:A14", nameRange: "D2");
            chart.AddSeries("E3:E14", "A3:A14", nameRange: "E2");
            chart.SetPosition(fromRow, fromCol, toRow, toCol);
            return chart;
        }

        private static void WriteStreamingWorkbook(string path, Random rng)
        {
            using ExcelStreamingWorkbook sw = ExcelStreamingWorkbook.Create(
                path,
                new ExcelStreamingOptions { UseInlineStrings = true });

            ExcelCellStyle header = sw.CreateStyle();
            header.Font.Bold = true;
            header.Font.Color = ExcelColor.White;
            header.Fill.Pattern = ExcelFillPattern.Solid;
            header.Fill.ForegroundColor = HeaderFill;

            ExcelCellStyle number = sw.CreateStyle();
            number.NumberFormat = NumberFormats.TwoDecimals;

            ExcelCellStyle date = sw.CreateStyle();
            date.NumberFormat = NumberFormats.DateTime;

            ExcelStreamingWorksheet sheet = sw.AddWorksheet("RandomRows");
            sheet.SetColumnWidth(1, 14);
            sheet.SetColumnWidth(2, 16);
            sheet.SetColumnWidth(3, 12);
            sheet.SetColumnWidth(4, 12);
            sheet.SetColumnWidth(5, 22);
            sheet.WriteRow(1, row =>
            {
                row.Text(1, "Station", header);
                row.Text(2, "Product", header);
                row.Text(3, "Units", header);
                row.Text(4, "Value", header);
                row.Text(5, "Stamp", header);
            });

            for (int i = 1; i <= 500; i++)
            {
                int n = i;
                sheet.WriteRow(n + 1, row =>
                {
                    row.Text(1, "ST-" + n.ToString("D4", CultureInfo.InvariantCulture));
                    row.Text(2, Products[n % Products.Length]);
                    row.Integer(3, rng.Next(1, 80));
                    row.Number(4, Math.Round(20 + rng.NextDouble() * 400, 2), number);
                    row.Date(5, DateTime.Now.AddMinutes(-n), date);
                });
            }

            sheet.Complete();

            ExcelStreamingWorksheet summary = sw.AddWorksheet("Repeats");
            summary.SetColumnWidth(1, 16);
            summary.WriteRow(1, new object?[] { "Region", "Flag" });
            for (int i = 0; i < 20; i++)
            {
                summary.WriteRow(i + 2, new object?[] { Regions[i % Regions.Length], i % 2 == 0 });
            }

            summary.Complete();
        }

        private static void VerifyRoundTrip(string path)
        {
            using ExcelWorkbook loaded = ExcelWorkbook.Open(path);
            Console.WriteLine("Opened " + Path.GetFileName(path));
            Console.WriteLine("  Sheets : " + loaded.Worksheets.Count.ToString(CultureInfo.InvariantCulture));
            int charts = 0;
            foreach (ExcelWorksheet sheet in loaded.Worksheets)
            {
                charts += sheet.Charts.Count;
                Console.WriteLine(
                    "    - " + sheet.Name
                    + "  cells=" + sheet.Cells.Count.ToString(CultureInfo.InvariantCulture)
                    + "  charts=" + sheet.Charts.Count.ToString(CultureInfo.InvariantCulture));
            }

            Console.WriteLine("  Charts : " + charts.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("  Names  : " + loaded.NamedRanges.Count.ToString(CultureInfo.InvariantCulture));
        }

        private static DataTable BuildOrdersTable(Random rng)
        {
            var table = new DataTable("Orders");
            table.Columns.Add("OrderId", typeof(string));
            table.Columns.Add("Product", typeof(string));
            table.Columns.Add("Qty", typeof(int));
            table.Columns.Add("Amount", typeof(double));
            table.Columns.Add("Shipped", typeof(DateTime));

            for (int i = 0; i < 10; i++)
            {
                DataRow row = table.NewRow();
                row["OrderId"] = "ORD-" + (1000 + i).ToString(CultureInfo.InvariantCulture);
                row["Product"] = Products[rng.Next(Products.Length)];
                row["Qty"] = rng.Next(1, 24);
                row["Amount"] = Math.Round(rng.Next(50, 900) + rng.NextDouble(), 2);
                row["Shipped"] = DateTime.Today.AddDays(-rng.Next(0, 21));
                table.Rows.Add(row);
            }

            return table;
        }

        private static void WriteHeaderRow(ExcelWorksheet sheet, int row, params string[] titles)
        {
            for (int i = 0; i < titles.Length; i++)
            {
                sheet[row, i + 1].SetValue(titles[i]);
            }

            StyleHeaderRow(sheet, row, titles.Length);
        }

        private static void StyleHeaderRow(ExcelWorksheet sheet, int row, int columns)
        {
            for (int c = 1; c <= columns; c++)
            {
                StyleHeaderCell(sheet[row, c]);
            }
        }

        private static void StyleHeaderCell(ExcelCell cell)
        {
            cell.Style.Font.Bold = true;
            cell.Style.Font.Color = ExcelColor.White;
            cell.Style.Fill.Pattern = ExcelFillPattern.Solid;
            cell.Style.Fill.ForegroundColor = HeaderFill;
            cell.Style.Alignment.Horizontal = ExcelHorizontalAlignment.Center;
            cell.Style.Borders.SetAll(ExcelBorderStyle.Thin, ExcelColor.FromRgb(0, 51, 102));
        }

        private static void ApplySectionStyle(ExcelCell cell)
        {
            cell.Style.Font.Bold = true;
            cell.Style.Font.Size = 12;
            cell.Style.Fill.Pattern = ExcelFillPattern.Solid;
            cell.Style.Fill.ForegroundColor = AccentFill;
        }

        private static void ApplyAltRow(ExcelWorksheet sheet, int row, int columns)
        {
            for (int c = 1; c <= columns; c++)
            {
                ExcelCell cell = sheet[row, c];
                if (cell.Style.Fill.Pattern == ExcelFillPattern.None)
                {
                    cell.Style.Fill.Pattern = ExcelFillPattern.Solid;
                    cell.Style.Fill.ForegroundColor = AltRowFill;
                }
            }
        }

        private static string SanitizeFileToken(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "demo";
            }

            char[] buffer = value.ToCharArray();
            for (int i = 0; i < buffer.Length; i++)
            {
                char ch = buffer[i];
                if (!(char.IsLetterOrDigit(ch) || ch == '.' || ch == '-' || ch == '_'))
                {
                    buffer[i] = '-';
                }
            }

            return new string(buffer).Trim('-');
        }
    }

    public sealed class StationReading
    {
        public string Station { get; set; } = string.Empty;
        public double Actual { get; set; }
        public double Target { get; set; }
        public bool InSpec { get; set; }
        public DateTime TakenAt { get; set; }
    }
}
