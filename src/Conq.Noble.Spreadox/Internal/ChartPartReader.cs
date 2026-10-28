using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Conq.Noble.Spreadox.Internal
{
    internal static class ChartPartReader
    {
        private static readonly Regex AbsoluteCellRegex = new Regex(
            @"\$?(?<col>[A-Za-z]+)\$?(?<row>\d+)",
            RegexOptions.Compiled);

        public static void LoadCharts(ExcelWorksheet worksheet, WorksheetPart worksheetPart)
        {
            DrawingsPart? drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart?.WorksheetDrawing is null)
            {
                return;
            }

            foreach (Xdr.TwoCellAnchor anchor in drawingsPart.WorksheetDrawing.Elements<Xdr.TwoCellAnchor>())
            {
                if (TryLoadPicture(worksheet, drawingsPart, anchor))
                {
                    continue;
                }

                Xdr.GraphicFrame? frame = anchor.GetFirstChild<Xdr.GraphicFrame>();
                A.GraphicData? graphicData = frame?
                    .GetFirstChild<A.Graphic>()?
                    .GetFirstChild<A.GraphicData>();

                C.ChartReference? chartRef = graphicData?.GetFirstChild<C.ChartReference>();
                string? relId = chartRef?.Id?.Value ?? ReadRelationshipId(graphicData);
                if (string.IsNullOrEmpty(relId))
                {
                    continue;
                }

                OpenXmlPart part = drawingsPart.GetPartById(relId!);
                if (part is ExtendedChartPart exPart)
                {
                    Cx.ChartSpace? exSpace = exPart.ChartSpace;
                    if (exSpace is null)
                    {
                        continue;
                    }

                    ExcelChart chart = worksheet.Charts.Add(ExcelChartType.BoxAndWhisker, ReadExTitle(exSpace));
                    ApplyAnchor(chart, anchor);
                    chart.ShowLegend = exSpace.GetFirstChild<Cx.Chart>()?.GetFirstChild<Cx.Legend>() != null;
                    LoadExSeries(chart, exSpace);
                    continue;
                }

                if (!(part is ChartPart chartPart) || chartPart.ChartSpace is null)
                {
                    continue;
                }

                ExcelChartType type = DetectType(chartPart.ChartSpace);
                string? title = ReadTitle(chartPart.ChartSpace);
                ExcelChart loaded = worksheet.Charts.Add(type, title);
                ApplyAnchor(loaded, anchor);
                ApplyLegend(loaded, chartPart.ChartSpace);
                LoadSeries(loaded, chartPart.ChartSpace);
            }
        }

        private static string? ReadRelationshipId(A.GraphicData? graphicData)
        {
            if (graphicData is null)
            {
                return null;
            }

            const string relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            foreach (OpenXmlElement child in graphicData.ChildElements)
            {
                OpenXmlAttribute attr = child.GetAttributes().FirstOrDefault(a =>
                    a.LocalName == "id" && (a.NamespaceUri == relNs || string.IsNullOrEmpty(a.NamespaceUri)));
                if (!string.IsNullOrEmpty(attr.Value))
                {
                    return attr.Value;
                }
            }

            return null;
        }

        private static bool TryLoadPicture(ExcelWorksheet worksheet, DrawingsPart drawingsPart, Xdr.TwoCellAnchor anchor)
        {
            Xdr.Picture? pic = anchor.GetFirstChild<Xdr.Picture>();
            if (pic is null)
            {
                return false;
            }

            string? embed = pic.GetFirstChild<Xdr.BlipFill>()?.GetFirstChild<A.Blip>()?.Embed?.Value;
            if (string.IsNullOrEmpty(embed) || !(drawingsPart.GetPartById(embed!) is ImagePart imagePart))
            {
                return true;
            }

            byte[] bytes;
            using (Stream stream = imagePart.GetStream())
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                bytes = ms.ToArray();
            }

            ExcelRange range = ReadAnchorRange(anchor);
            bool lockAspect = pic.GetFirstChild<Xdr.NonVisualPictureProperties>()?
                .GetFirstChild<Xdr.NonVisualPictureDrawingProperties>()?
                .GetFirstChild<A.PictureLocks>()?
                .NoChangeAspect?.Value ?? false;

            ExcelPicture picture = worksheet.Pictures.Add(bytes, range.A1, lockAspect);
            picture.Name = pic.GetFirstChild<Xdr.NonVisualPictureProperties>()?
                .GetFirstChild<Xdr.NonVisualDrawingProperties>()?
                .Name?.Value;
            return true;
        }

        private static ExcelRange ReadAnchorRange(Xdr.TwoCellAnchor anchor)
        {
            Xdr.FromMarker? from = anchor.FromMarker;
            Xdr.ToMarker? to = anchor.ToMarker;
            int fromCol = ParseMarkerIndex(from?.ColumnId?.Text) + 1;
            int fromRow = ParseMarkerIndex(from?.RowId?.Text) + 1;
            int toCol = ParseMarkerIndex(to?.ColumnId?.Text) + 1;
            int toRow = ParseMarkerIndex(to?.RowId?.Text) + 1;
            if (fromCol < 1) fromCol = 1;
            if (fromRow < 1) fromRow = 1;
            if (toCol < fromCol) toCol = fromCol;
            if (toRow < fromRow) toRow = fromRow;
            return new ExcelRange(fromRow, fromCol, toRow, toCol);
        }

        private static string? ReadExTitle(Cx.ChartSpace space)
        {
            Cx.ChartTitle? title = space.GetFirstChild<Cx.Chart>()?.GetFirstChild<Cx.ChartTitle>();
            Cx.VXsdstring? text = title?.Descendants<Cx.VXsdstring>().FirstOrDefault();
            return text?.Text;
        }

        private static void LoadExSeries(ExcelChart chart, Cx.ChartSpace space)
        {
            Cx.Series? series = space.Descendants<Cx.Series>().FirstOrDefault();
            if (series is null)
            {
                return;
            }

            string? name = series.GetFirstChild<Cx.Text>()?.Descendants<Cx.VXsdstring>().FirstOrDefault()?.Text;
            Cx.Formula? valFormula = space.Descendants<Cx.NumericDimension>().FirstOrDefault()?.GetFirstChild<Cx.Formula>();
            Cx.Formula? catFormula = space.Descendants<Cx.StringDimension>().FirstOrDefault()?.GetFirstChild<Cx.Formula>();
            chart.AddSeries(
                StripSheet(valFormula?.Text) ?? "A2:A2",
                StripSheet(catFormula?.Text),
                name: name);

            Cx.SeriesElementVisibilities? vis = series.Descendants<Cx.SeriesElementVisibilities>().FirstOrDefault();
            if (vis != null)
            {
                chart.BoxAndWhisker.ShowOutlierPoints = vis.Outliers?.Value ?? true;
                chart.BoxAndWhisker.ShowInnerPoints = vis.Nonoutliers?.Value ?? true;
                chart.BoxAndWhisker.ShowMeanMarkers = vis.MeanMarker?.Value ?? true;
                chart.BoxAndWhisker.ShowMeanLine = vis.MeanLine?.Value ?? false;
            }
        }

        private static ExcelChartType DetectType(C.ChartSpace space)
        {
            C.PlotArea? plot = space.GetFirstChild<C.Chart>()?.GetFirstChild<C.PlotArea>();
            if (plot is null)
            {
                return ExcelChartType.ColumnClustered;
            }

            int distinct = 0;
            if (plot.GetFirstChild<C.BarChart>() != null) distinct++;
            if (plot.GetFirstChild<C.LineChart>() != null) distinct++;
            if (plot.GetFirstChild<C.AreaChart>() != null) distinct++;
            if (plot.GetFirstChild<C.ScatterChart>() != null) distinct++;
            if (plot.GetFirstChild<C.PieChart>() != null) distinct++;
            if (plot.GetFirstChild<C.DoughnutChart>() != null) distinct++;
            if (plot.GetFirstChild<C.BubbleChart>() != null) distinct++;
            if (plot.GetFirstChild<C.RadarChart>() != null) distinct++;
            if (distinct > 1)
            {
                return ExcelChartType.Combo;
            }

            if (plot.GetFirstChild<C.PieChart>() != null) return ExcelChartType.Pie;
            if (plot.GetFirstChild<C.DoughnutChart>() != null) return ExcelChartType.Doughnut;
            if (plot.GetFirstChild<C.BubbleChart>() != null) return ExcelChartType.Bubble;
            if (plot.GetFirstChild<C.RadarChart>() != null) return ExcelChartType.Radar;
            if (plot.GetFirstChild<C.ScatterChart>() != null) return ExcelChartType.Scatter;

            C.AreaChart? area = plot.GetFirstChild<C.AreaChart>();
            if (area != null)
            {
                return FromAreaGrouping(area.Grouping?.Val?.Value);
            }

            C.LineChart? line = plot.GetFirstChild<C.LineChart>();
            if (line != null)
            {
                return FromLineGrouping(line.Grouping?.Val?.Value);
            }

            C.BarChart? bar = plot.GetFirstChild<C.BarChart>();
            if (bar != null)
            {
                bool isColumn = bar.BarDirection?.Val?.Value == C.BarDirectionValues.Column;
                return FromBarGrouping(bar.BarGrouping?.Val?.Value, isColumn);
            }

            return ExcelChartType.ColumnClustered;
        }

        private static ExcelChartType FromBarGrouping(C.BarGroupingValues? grouping, bool column)
        {
            if (grouping == C.BarGroupingValues.Stacked)
            {
                return column ? ExcelChartType.ColumnStacked : ExcelChartType.BarStacked;
            }

            if (grouping == C.BarGroupingValues.PercentStacked)
            {
                return column ? ExcelChartType.ColumnStacked100 : ExcelChartType.BarStacked100;
            }

            return column ? ExcelChartType.ColumnClustered : ExcelChartType.BarClustered;
        }

        private static ExcelChartType FromLineGrouping(C.GroupingValues? grouping)
        {
            if (grouping == C.GroupingValues.Stacked) return ExcelChartType.LineStacked;
            if (grouping == C.GroupingValues.PercentStacked) return ExcelChartType.LineStacked100;
            return ExcelChartType.Line;
        }

        private static ExcelChartType FromAreaGrouping(C.GroupingValues? grouping)
        {
            if (grouping == C.GroupingValues.Stacked) return ExcelChartType.AreaStacked;
            if (grouping == C.GroupingValues.PercentStacked) return ExcelChartType.AreaStacked100;
            return ExcelChartType.Area;
        }

        private static string? ReadTitle(C.ChartSpace space)
        {
            C.Title? title = space.GetFirstChild<C.Chart>()?.GetFirstChild<C.Title>();
            A.Text? text = title?.Descendants<A.Text>().FirstOrDefault();
            return text?.Text;
        }

        private static void ApplyLegend(ExcelChart chart, C.ChartSpace space)
        {
            C.Legend? legend = space.GetFirstChild<C.Chart>()?.GetFirstChild<C.Legend>();
            if (legend is null)
            {
                chart.ShowLegend = false;
                return;
            }

            chart.ShowLegend = true;
            C.LegendPositionValues? pos = legend.LegendPosition?.Val?.Value;
            if (pos == C.LegendPositionValues.Left) chart.LegendPosition = ExcelLegendPosition.Left;
            else if (pos == C.LegendPositionValues.Top) chart.LegendPosition = ExcelLegendPosition.Top;
            else if (pos == C.LegendPositionValues.Bottom) chart.LegendPosition = ExcelLegendPosition.Bottom;
            else if (pos == C.LegendPositionValues.TopRight) chart.LegendPosition = ExcelLegendPosition.TopRight;
            else chart.LegendPosition = ExcelLegendPosition.Right;
        }

        private static void ApplyAnchor(ExcelChart chart, Xdr.TwoCellAnchor anchor)
        {
            Xdr.FromMarker? from = anchor.FromMarker;
            Xdr.ToMarker? to = anchor.ToMarker;
            if (from is null || to is null)
            {
                return;
            }

            int fromCol = ParseMarkerIndex(from.ColumnId?.Text) + 1;
            int fromRow = ParseMarkerIndex(from.RowId?.Text) + 1;
            int toCol = ParseMarkerIndex(to.ColumnId?.Text) + 1;
            int toRow = ParseMarkerIndex(to.RowId?.Text) + 1;
            if (fromCol > 0 && fromRow > 0 && toCol >= fromCol && toRow >= fromRow)
            {
                chart.SetPosition(fromRow, fromCol, toRow, toCol);
            }
        }

        private static int ParseMarkerIndex(string? text)
        {
            if (int.TryParse(text, out int value) && value >= 0)
            {
                return value;
            }

            return 0;
        }

        private static void LoadSeries(ExcelChart chart, C.ChartSpace space)
        {
            C.PlotArea? plot = space.GetFirstChild<C.Chart>()?.GetFirstChild<C.PlotArea>();
            if (plot is null)
            {
                return;
            }

            foreach (C.BarChart bar in plot.Elements<C.BarChart>())
            {
                ExcelChartType seriesType = FromBarGrouping(
                    bar.BarGrouping?.Val?.Value,
                    bar.BarDirection?.Val?.Value == C.BarDirectionValues.Column);
                bool secondary = UsesSecondary(bar);
                foreach (C.BarChartSeries ser in bar.Elements<C.BarChartSeries>())
                {
                    AddCategoryValueSeries(chart, ser.GetFirstChild<C.SeriesText>(), ser.GetFirstChild<C.CategoryAxisData>(), ser.GetFirstChild<C.Values>(), seriesType, secondary);
                }
            }

            foreach (C.LineChart line in plot.Elements<C.LineChart>())
            {
                ExcelChartType seriesType = FromLineGrouping(line.Grouping?.Val?.Value);
                bool secondary = UsesSecondary(line);
                foreach (C.LineChartSeries ser in line.Elements<C.LineChartSeries>())
                {
                    AddCategoryValueSeries(chart, ser.GetFirstChild<C.SeriesText>(), ser.GetFirstChild<C.CategoryAxisData>(), ser.GetFirstChild<C.Values>(), seriesType, secondary);
                }
            }

            foreach (C.AreaChart area in plot.Elements<C.AreaChart>())
            {
                ExcelChartType seriesType = FromAreaGrouping(area.Grouping?.Val?.Value);
                bool secondary = UsesSecondary(area);
                foreach (C.AreaChartSeries ser in area.Elements<C.AreaChartSeries>())
                {
                    AddCategoryValueSeries(chart, ser.GetFirstChild<C.SeriesText>(), ser.GetFirstChild<C.CategoryAxisData>(), ser.GetFirstChild<C.Values>(), seriesType, secondary);
                }
            }

            foreach (C.ScatterChart scatter in plot.Elements<C.ScatterChart>())
            {
                foreach (C.ScatterChartSeries ser in scatter.Elements<C.ScatterChartSeries>())
                {
                    string? name = ReadSeriesName(ser.GetFirstChild<C.SeriesText>());
                    string? x = ReadFormula(ser.GetFirstChild<C.XValues>()?.GetFirstChild<C.NumberReference>()?.Formula);
                    string? y = ReadFormula(ser.GetFirstChild<C.YValues>()?.GetFirstChild<C.NumberReference>()?.Formula);
                    if (!string.IsNullOrWhiteSpace(y))
                    {
                        ExcelChartSeries loaded = chart.AddSeries(StripSheet(y!)!, StripSheet(x), name);
                        if (chart.Type == ExcelChartType.Combo)
                        {
                            loaded.SeriesChartType = ExcelChartType.Scatter;
                        }
                    }
                }
            }

            foreach (C.PieChart pie in plot.Elements<C.PieChart>())
            {
                foreach (C.PieChartSeries ser in pie.Elements<C.PieChartSeries>())
                {
                    AddCategoryValueSeries(chart, ser.GetFirstChild<C.SeriesText>(), ser.GetFirstChild<C.CategoryAxisData>(), ser.GetFirstChild<C.Values>(), ExcelChartType.Pie, false);
                }
            }

            foreach (C.DoughnutChart donut in plot.Elements<C.DoughnutChart>())
            {
                C.HoleSize? hole = donut.GetFirstChild<C.HoleSize>();
                if (hole?.Val?.HasValue == true)
                {
                    chart.DoughnutHoleSize = hole.Val.Value;
                }

                foreach (C.PieChartSeries ser in donut.Elements<C.PieChartSeries>())
                {
                    AddCategoryValueSeries(chart, ser.GetFirstChild<C.SeriesText>(), ser.GetFirstChild<C.CategoryAxisData>(), ser.GetFirstChild<C.Values>(), ExcelChartType.Doughnut, false);
                }
            }

            foreach (C.BubbleChart bubble in plot.Elements<C.BubbleChart>())
            {
                foreach (C.BubbleChartSeries ser in bubble.Elements<C.BubbleChartSeries>())
                {
                    string? name = ReadSeriesName(ser.GetFirstChild<C.SeriesText>());
                    string? x = ReadFormula(ser.GetFirstChild<C.XValues>()?.GetFirstChild<C.NumberReference>()?.Formula);
                    string? y = ReadFormula(ser.GetFirstChild<C.YValues>()?.GetFirstChild<C.NumberReference>()?.Formula);
                    string? size = ReadFormula(ser.GetFirstChild<C.BubbleSize>()?.GetFirstChild<C.NumberReference>()?.Formula);
                    if (!string.IsNullOrWhiteSpace(y))
                    {
                        ExcelChartSeries loaded = chart.AddSeries(StripSheet(y!)!, StripSheet(x), name);
                        loaded.BubbleSizeRange = StripSheet(size);
                        if (chart.Type == ExcelChartType.Combo)
                        {
                            loaded.SeriesChartType = ExcelChartType.Bubble;
                        }
                    }
                }
            }

            foreach (C.RadarChart radar in plot.Elements<C.RadarChart>())
            {
                bool secondary = UsesSecondary(radar);
                foreach (C.RadarChartSeries ser in radar.Elements<C.RadarChartSeries>())
                {
                    AddCategoryValueSeries(chart, ser.GetFirstChild<C.SeriesText>(), ser.GetFirstChild<C.CategoryAxisData>(), ser.GetFirstChild<C.Values>(), ExcelChartType.Radar, secondary);
                }
            }
        }

        private static bool UsesSecondary(OpenXmlElement chartElement)
        {
            foreach (C.AxisId ax in chartElement.Elements<C.AxisId>())
            {
                if (ax.Val?.Value == 3)
                {
                    return true;
                }
            }

            return false;
        }

        private static void AddCategoryValueSeries(
            ExcelChart chart,
            C.SeriesText? seriesText,
            C.CategoryAxisData? cats,
            C.Values? values,
            ExcelChartType seriesType,
            bool secondary)
        {
            string? name = ReadSeriesName(seriesText);
            string? catFormula = ReadFormula(cats?.GetFirstChild<C.StringReference>()?.Formula
                                            ?? cats?.GetFirstChild<C.NumberReference>()?.Formula);
            string? valFormula = ReadFormula(values?.GetFirstChild<C.NumberReference>()?.Formula);
            if (string.IsNullOrWhiteSpace(valFormula))
            {
                return;
            }

            ExcelChartSeries loaded = chart.AddSeries(StripSheet(valFormula!)!, StripSheet(catFormula), name);
            loaded.UseSecondaryAxis = secondary;
            if (chart.Type == ExcelChartType.Combo)
            {
                loaded.SeriesChartType = seriesType;
            }
        }

        private static string? ReadSeriesName(C.SeriesText? text)
        {
            if (text is null)
            {
                return null;
            }

            string? literal = text.NumericValue?.Text;
            if (!string.IsNullOrEmpty(literal))
            {
                return literal;
            }

            return text.StringReference?.Formula?.Text;
        }

        private static string? ReadFormula(C.Formula? formula) => formula?.Text;

        private static string? StripSheet(string? formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return null;
            }

            string text = formula!.Trim();
            int bang = text.LastIndexOf('!');
            if (bang >= 0 && bang < text.Length - 1)
            {
                text = text.Substring(bang + 1);
            }

            MatchCollection matches = AbsoluteCellRegex.Matches(text);
            if (matches.Count == 1)
            {
                return matches[0].Groups["col"].Value.ToUpperInvariant() + matches[0].Groups["row"].Value;
            }

            if (matches.Count >= 2)
            {
                Match a = matches[0];
                Match b = matches[matches.Count - 1];
                return a.Groups["col"].Value.ToUpperInvariant() + a.Groups["row"].Value
                       + ":"
                       + b.Groups["col"].Value.ToUpperInvariant() + b.Groups["row"].Value;
            }

            return text.Replace("$", string.Empty);
        }
    }
}
