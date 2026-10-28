using System;
using System.Collections.Generic;
using PdfSharpCore.Drawing;

namespace Conq.Noble.Spreadox.Internal
{
    internal static class PdfChartPainter
    {
        private static readonly ExcelColor[] Palette =
        {
            ExcelColor.FromRgb(0, 82, 147),
            ExcelColor.FromRgb(237, 125, 49),
            ExcelColor.FromRgb(91, 155, 213),
            ExcelColor.FromRgb(112, 173, 71),
            ExcelColor.FromRgb(255, 192, 0),
            ExcelColor.FromRgb(165, 165, 165),
            ExcelColor.FromRgb(158, 72, 14),
            ExcelColor.FromRgb(68, 114, 196)
        };

        private enum PlotKind
        {
            Line,
            LineMarker,
            Area,
            Bar,
            Pie,
            Combo
        }

        public static void Draw(XGraphics gfx, XRect bounds, PdfChart chart)
        {
            if (!HasPoints(chart))
            {
                DrawCentered(gfx, bounds, "No chart data", 12, XFontStyle.Italic);
                return;
            }

            PlotKind kind = Resolve(chart.Type);
            if (kind == PlotKind.Pie)
            {
                DrawPie(gfx, bounds, chart);
                return;
            }

            double legendHeight = chart.ShowLegend ? 22 : 0;
            var plot = new XRect(bounds.X + 42, bounds.Y + 8, bounds.Width - 78, bounds.Height - 56 - legendHeight);
            if (plot.Width < 40 || plot.Height < 40)
            {
                return;
            }

            IReadOnlyList<string> categories = chart.Series[0].Categories;
            int n = categories.Count;
            GetPrimaryRange(chart, out double yMin, out double yMax);
            GetSecondaryRange(chart, out double y2Min, out double y2Max);

            DrawAxes(gfx, plot, chart, categories, yMin, yMax, y2Min, y2Max, kind == PlotKind.Combo && HasSecondary(chart));

            if (kind == PlotKind.Combo)
            {
                DrawBars(gfx, plot, chart, primaryOnly: true, yMin, yMax, n);
                DrawLines(gfx, plot, chart, secondaryOnly: true, y2Min, y2Max, n, markers: true);
            }
            else if (kind == PlotKind.Bar)
            {
                DrawBars(gfx, plot, chart, primaryOnly: false, yMin, yMax, n);
            }
            else if (kind == PlotKind.Area)
            {
                DrawArea(gfx, plot, chart, yMin, yMax, n);
            }
            else
            {
                DrawLines(gfx, plot, chart, secondaryOnly: false, yMin, yMax, n, markers: kind == PlotKind.LineMarker);
            }

            DrawConstantLines(gfx, plot, chart, yMin, yMax);
            if (chart.ShowLegend)
            {
                DrawLegend(gfx, new XRect(bounds.X, bounds.Bottom - legendHeight, bounds.Width, legendHeight), chart);
            }
        }

        private static PlotKind Resolve(ExcelChartType type)
        {
            switch (type)
            {
                case ExcelChartType.Pie:
                case ExcelChartType.Doughnut:
                    return PlotKind.Pie;
                case ExcelChartType.Combo:
                    return PlotKind.Combo;
                case ExcelChartType.Area:
                case ExcelChartType.AreaStacked:
                case ExcelChartType.AreaStacked100:
                    return PlotKind.Area;
                case ExcelChartType.LineMarker:
                case ExcelChartType.BoxAndWhisker:
                    return PlotKind.LineMarker;
                case ExcelChartType.Line:
                case ExcelChartType.LineStacked:
                case ExcelChartType.LineStacked100:
                    return PlotKind.Line;
                default:
                    return PlotKind.Bar;
            }
        }

        private static bool HasPoints(PdfChart chart)
        {
            if (chart.Series.Count == 0)
            {
                return false;
            }

            for (int i = 0; i < chart.Series.Count; i++)
            {
                if (chart.Series[i].Values.Count > 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool HasSecondary(PdfChart chart)
        {
            for (int i = 0; i < chart.Series.Count; i++)
            {
                if (chart.Series[i].UseSecondaryAxis)
                {
                    return true;
                }
            }

            return false;
        }

        private static ExcelColor SeriesColor(PdfChartSeries series, int index)
        {
            return series.Color ?? Palette[index % Palette.Length];
        }

        private static XColor ToX(ExcelColor color)
        {
            return XColor.FromArgb(color.R, color.G, color.B);
        }

        private static void GetPrimaryRange(PdfChart chart, out double min, out double max)
        {
            min = double.PositiveInfinity;
            max = double.NegativeInfinity;
            for (int s = 0; s < chart.Series.Count; s++)
            {
                PdfChartSeries series = chart.Series[s];
                if (series.UseSecondaryAxis)
                {
                    continue;
                }

                for (int i = 0; i < series.Values.Count; i++)
                {
                    double v = series.Values[i];
                    if (double.IsNaN(v) || double.IsInfinity(v))
                    {
                        continue;
                    }

                    if (v < min) min = v;
                    if (v > max) max = v;
                }
            }

            for (int i = 0; i < chart.ConstantLines.Count; i++)
            {
                double v = chart.ConstantLines[i].Value;
                if (v < min) min = v;
                if (v > max) max = v;
            }

            if (chart.ValueAxis.Minimum.HasValue) min = chart.ValueAxis.Minimum.Value;
            if (chart.ValueAxis.Maximum.HasValue) max = chart.ValueAxis.Maximum.Value;
            PadRange(ref min, ref max);
        }

        private static void GetSecondaryRange(PdfChart chart, out double min, out double max)
        {
            min = chart.SecondaryValueAxis.Minimum ?? 0;
            max = chart.SecondaryValueAxis.Maximum ?? 100;
            if (max <= min)
            {
                max = min + 1;
            }
        }

        private static void PadRange(ref double min, ref double max)
        {
            if (double.IsInfinity(min) || double.IsInfinity(max))
            {
                min = 0;
                max = 1;
                return;
            }

            if (Math.Abs(max - min) < 1e-9)
            {
                min -= 1;
                max += 1;
            }

            double pad = (max - min) * 0.08;
            min -= pad;
            max += pad;
            if (min > 0 && min - pad < 0)
            {
                min = 0;
            }
        }

        private static double MapX(XRect plot, int index, int n)
        {
            if (n <= 1)
            {
                return plot.X + plot.Width / 2;
            }

            return plot.X + plot.Width * index / (n - 1);
        }

        private static double MapY(XRect plot, double value, double min, double max)
        {
            double t = (value - min) / (max - min);
            return plot.Y + plot.Height * (1 - t);
        }

        private static void DrawAxes(
            XGraphics gfx,
            XRect plot,
            PdfChart chart,
            IReadOnlyList<string> categories,
            double yMin,
            double yMax,
            double y2Min,
            double y2Max,
            bool secondary)
        {
            var axisPen = new XPen(XColors.Gray, 0.6);
            var gridPen = new XPen(XColor.FromArgb(220, 220, 220), 0.5);
            gfx.DrawRectangle(axisPen, plot);

            var tickFont = new XFont(PdfFontResolver.FamilyName, 7, XFontStyle.Regular);
            var titleFont = new XFont(PdfFontResolver.FamilyName, 8, XFontStyle.Regular);
            var brush = new XSolidBrush(XColors.DimGray);

            int ticks = 5;
            for (int i = 0; i <= ticks; i++)
            {
                double value = yMin + (yMax - yMin) * i / ticks;
                double y = MapY(plot, value, yMin, yMax);
                if (chart.ValueAxis.ShowMajorGridlines && i > 0 && i < ticks)
                {
                    gfx.DrawLine(gridPen, plot.X, y, plot.Right, y);
                }

                string label = FormatTick(value);
                gfx.DrawString(label, tickFont, brush, new XRect(plot.X - 40, y - 6, 36, 12), XStringFormats.CenterRight);
            }

            if (secondary)
            {
                for (int i = 0; i <= ticks; i++)
                {
                    double value = y2Min + (y2Max - y2Min) * i / ticks;
                    double y = MapY(plot, value, y2Min, y2Max);
                    string label = FormatTick(value);
                    gfx.DrawString(label, tickFont, brush, new XRect(plot.Right + 4, y - 6, 30, 12), XStringFormats.CenterLeft);
                }

                if (!string.IsNullOrWhiteSpace(chart.SecondaryValueAxis.Title))
                {
                    gfx.DrawString(chart.SecondaryValueAxis.Title, titleFont, brush,
                        new XRect(plot.Right + 4, plot.Y - 12, 30, 12), XStringFormats.CenterLeft);
                }
            }

            int n = categories.Count;
            int step = n <= 16 ? 1 : Math.Max(1, n / 12);
            double rotation = chart.CategoryAxis.LabelRotation;
            for (int i = 0; i < n; i += step)
            {
                double x = MapX(plot, i, n);
                gfx.DrawLine(axisPen, x, plot.Bottom, x, plot.Bottom + 3);
                DrawRotated(gfx, categories[i] ?? string.Empty, tickFont, brush, x, plot.Bottom + 5, rotation);
            }

            if (!string.IsNullOrWhiteSpace(chart.ValueAxis.Title))
            {
                gfx.DrawString(chart.ValueAxis.Title, titleFont, brush,
                    new XRect(plot.X - 42, plot.Y - 14, 80, 12), XStringFormats.TopLeft);
            }
        }

        private static void DrawRotated(XGraphics gfx, string text, XFont font, XBrush brush, double x, double y, double degrees)
        {
            if (Math.Abs(degrees) < 0.5)
            {
                gfx.DrawString(text, font, brush, new XRect(x - 30, y, 60, 12), XStringFormats.TopCenter);
                return;
            }

            gfx.Save();
            gfx.TranslateTransform(x, y);
            gfx.RotateTransform(degrees);
            gfx.DrawString(text, font, brush, new XRect(-40, 0, 40, 12), XStringFormats.TopRight);
            gfx.Restore();
        }

        private static void DrawBars(XGraphics gfx, XRect plot, PdfChart chart, bool primaryOnly, double yMin, double yMax, int n)
        {
            var barSeries = new List<PdfChartSeries>();
            var colors = new List<ExcelColor>();
            for (int i = 0; i < chart.Series.Count; i++)
            {
                PdfChartSeries series = chart.Series[i];
                bool isBar = IsBarSeries(chart, series);
                if (primaryOnly && (series.UseSecondaryAxis || !isBar))
                {
                    continue;
                }

                if (!primaryOnly && series.UseSecondaryAxis)
                {
                    continue;
                }

                barSeries.Add(series);
                colors.Add(SeriesColor(series, i));
            }

            if (barSeries.Count == 0 || n == 0)
            {
                return;
            }

            double slot = n <= 1 ? plot.Width * 0.5 : plot.Width / n;
            double cluster = slot * 0.72;
            double barW = cluster / barSeries.Count;
            double zeroY = MapY(plot, Clamp(0, yMin, yMax), yMin, yMax);

            for (int s = 0; s < barSeries.Count; s++)
            {
                var fill = new XSolidBrush(ToX(colors[s]));
                var pen = new XPen(ToX(colors[s]), 0.4);
                IReadOnlyList<double> values = barSeries[s].Values;
                for (int i = 0; i < n && i < values.Count; i++)
                {
                    double v = values[i];
                    if (double.IsNaN(v))
                    {
                        continue;
                    }

                    double center = n <= 1 ? plot.X + plot.Width / 2 : plot.X + slot * (i + 0.5);
                    double left = center - cluster / 2 + s * barW;
                    double y = MapY(plot, v, yMin, yMax);
                    double top = Math.Min(y, zeroY);
                    double height = Math.Abs(y - zeroY);
                    if (height < 0.5) height = 0.5;
                    gfx.DrawRectangle(pen, fill, left, top, Math.Max(barW - 1, 1), height);
                }
            }
        }

        private static void DrawLines(
            XGraphics gfx,
            XRect plot,
            PdfChart chart,
            bool secondaryOnly,
            double yMin,
            double yMax,
            int n,
            bool markers)
        {
            for (int s = 0; s < chart.Series.Count; s++)
            {
                PdfChartSeries series = chart.Series[s];
                if (secondaryOnly && !series.UseSecondaryAxis && !IsLineSeries(chart, series))
                {
                    continue;
                }

                if (secondaryOnly && !series.UseSecondaryAxis && IsBarSeries(chart, series))
                {
                    continue;
                }

                if (!secondaryOnly && series.UseSecondaryAxis)
                {
                    continue;
                }

                ExcelColor color = SeriesColor(series, s);
                var pen = new XPen(ToX(color), 2);
                IReadOnlyList<double> values = series.Values;
                XPoint? prev = null;
                for (int i = 0; i < n && i < values.Count; i++)
                {
                    double v = values[i];
                    if (double.IsNaN(v))
                    {
                        prev = null;
                        continue;
                    }

                    var pt = new XPoint(MapX(plot, i, n), MapY(plot, v, yMin, yMax));
                    if (prev.HasValue)
                    {
                        gfx.DrawLine(pen, prev.Value, pt);
                    }

                    if (markers || series.ShowMarkers)
                    {
                        gfx.DrawEllipse(new XSolidBrush(ToX(color)), pt.X - 3, pt.Y - 3, 6, 6);
                    }

                    prev = pt;
                }
            }
        }

        private static void DrawArea(XGraphics gfx, XRect plot, PdfChart chart, double yMin, double yMax, int n)
        {
            for (int s = 0; s < chart.Series.Count; s++)
            {
                PdfChartSeries series = chart.Series[s];
                if (series.UseSecondaryAxis)
                {
                    continue;
                }

                ExcelColor color = SeriesColor(series, s);
                XColor fillColor = XColor.FromArgb(90, color.R, color.G, color.B);
                var points = new List<XPoint>();
                double baseY = MapY(plot, Math.Max(yMin, 0), yMin, yMax);
                IReadOnlyList<double> values = series.Values;
                for (int i = 0; i < n && i < values.Count; i++)
                {
                    double v = values[i];
                    if (double.IsNaN(v))
                    {
                        v = yMin;
                    }

                    points.Add(new XPoint(MapX(plot, i, n), MapY(plot, v, yMin, yMax)));
                }

                if (points.Count == 0)
                {
                    continue;
                }

                var poly = new XPoint[points.Count + 2];
                poly[0] = new XPoint(points[0].X, baseY);
                for (int i = 0; i < points.Count; i++)
                {
                    poly[i + 1] = points[i];
                }

                poly[poly.Length - 1] = new XPoint(points[points.Count - 1].X, baseY);
                gfx.DrawPolygon(new XPen(ToX(color), 1.5), new XSolidBrush(fillColor), poly, XFillMode.Winding);
                for (int i = 0; i < points.Count; i++)
                {
                    gfx.DrawEllipse(new XSolidBrush(ToX(color)), points[i].X - 3, points[i].Y - 3, 6, 6);
                }
            }
        }

        private static void DrawPie(XGraphics gfx, XRect bounds, PdfChart chart)
        {
            PdfChartSeries series = chart.Series[0];
            double total = 0;
            for (int i = 0; i < series.Values.Count; i++)
            {
                if (series.Values[i] > 0)
                {
                    total += series.Values[i];
                }
            }

            if (total <= 0)
            {
                DrawCentered(gfx, bounds, "No chart data", 12, XFontStyle.Italic);
                return;
            }

            double size = Math.Min(bounds.Width, bounds.Height) * 0.62;
            var rect = new XRect(bounds.X + (bounds.Width - size) / 2 - 30, bounds.Y + (bounds.Height - size) / 2, size, size);
            double angle = -90;
            var labelFont = new XFont(PdfFontResolver.FamilyName, 7, XFontStyle.Regular);
            for (int i = 0; i < series.Values.Count; i++)
            {
                double v = series.Values[i];
                if (v <= 0)
                {
                    continue;
                }

                double sweep = 360.0 * v / total;
                ExcelColor color = Palette[i % Palette.Length];
                gfx.DrawPie(new XPen(XColors.White, 0.8), new XSolidBrush(ToX(color)), rect, angle, sweep);

                double mid = (angle + sweep / 2) * Math.PI / 180.0;
                double cx = rect.X + rect.Width / 2 + Math.Cos(mid) * (size / 2 + 18);
                double cy = rect.Y + rect.Height / 2 + Math.Sin(mid) * (size / 2 + 18);
                int pct = (int)Math.Round(100.0 * v / total);
                string cat = i < series.Categories.Count ? series.Categories[i] : (i + 1).ToString();
                gfx.DrawString(cat + ": " + pct + "%", labelFont, XBrushes.DimGray,
                    new XRect(cx - 40, cy - 6, 80, 12), XStringFormats.Center);

                angle += sweep;
            }

            if (chart.ShowLegend)
            {
                DrawLegend(gfx, new XRect(bounds.X, bounds.Bottom - 20, bounds.Width, 18), chart);
            }
        }

        private static void DrawConstantLines(XGraphics gfx, XRect plot, PdfChart chart, double yMin, double yMax)
        {
            var font = new XFont(PdfFontResolver.FamilyName, 7, XFontStyle.Bold);
            for (int i = 0; i < chart.ConstantLines.Count; i++)
            {
                PdfConstantLine line = chart.ConstantLines[i];
                if (line.Value < yMin || line.Value > yMax)
                {
                    continue;
                }

                double y = MapY(plot, line.Value, yMin, yMax);
                var pen = new XPen(ToX(line.Color), 1.2) { DashStyle = XDashStyle.Dash };
                gfx.DrawLine(pen, plot.X, y, plot.Right, y);
                gfx.DrawString(line.Name, font, new XSolidBrush(ToX(line.Color)),
                    new XRect(plot.Right - 80, y - 10, 78, 10), XStringFormats.CenterRight);
            }
        }

        private static void DrawLegend(XGraphics gfx, XRect bounds, PdfChart chart)
        {
            var font = new XFont(PdfFontResolver.FamilyName, 7, XFontStyle.Regular);
            double x = bounds.X + 8;
            double y = bounds.Y + 4;
            for (int i = 0; i < chart.Series.Count; i++)
            {
                PdfChartSeries series = chart.Series[i];
                string name = string.IsNullOrWhiteSpace(series.Name) ? "Series " + (i + 1) : series.Name!;
                ExcelColor color = SeriesColor(series, i);
                gfx.DrawRectangle(new XSolidBrush(ToX(color)), x, y, 8, 8);
                gfx.DrawString(name, font, XBrushes.DimGray, new XRect(x + 11, y - 2, 90, 12), XStringFormats.TopLeft);
                x += 110;
                if (x > bounds.Right - 80)
                {
                    break;
                }
            }

            for (int i = 0; i < chart.ConstantLines.Count; i++)
            {
                PdfConstantLine line = chart.ConstantLines[i];
                gfx.DrawRectangle(new XSolidBrush(ToX(line.Color)), x, y, 8, 8);
                gfx.DrawString(line.Name, font, XBrushes.DimGray, new XRect(x + 11, y - 2, 70, 12), XStringFormats.TopLeft);
                x += 90;
            }
        }

        private static ExcelChartType EffectiveType(PdfChart chart, PdfChartSeries series)
        {
            if (series.SeriesChartType.HasValue)
            {
                return series.SeriesChartType.Value;
            }

            if (chart.Type == ExcelChartType.Combo)
            {
                return series.UseSecondaryAxis ? ExcelChartType.Line : ExcelChartType.ColumnClustered;
            }

            return chart.Type;
        }

        private static bool IsBarSeries(PdfChart chart, PdfChartSeries series)
        {
            return Resolve(EffectiveType(chart, series)) == PlotKind.Bar;
        }

        private static bool IsLineSeries(PdfChart chart, PdfChartSeries series)
        {
            PlotKind kind = Resolve(EffectiveType(chart, series));
            return kind == PlotKind.Line || kind == PlotKind.LineMarker;
        }

        private static double Clamp(double value, double min, double max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        private static string FormatTick(double value)
        {
            double abs = Math.Abs(value);
            if (abs >= 1000)
            {
                return value.ToString("0");
            }

            if (abs >= 10)
            {
                return value.ToString("0.#");
            }

            return value.ToString("0.##");
        }

        private static void DrawCentered(XGraphics gfx, XRect bounds, string text, double size, XFontStyle style)
        {
            var font = new XFont(PdfFontResolver.FamilyName, size, style);
            gfx.DrawString(text, font, XBrushes.Gray, bounds, XStringFormats.Center);
        }
    }
}
