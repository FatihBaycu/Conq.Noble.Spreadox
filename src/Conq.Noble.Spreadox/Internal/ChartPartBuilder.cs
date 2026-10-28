using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Conq.Noble.Spreadox.Internal
{
    /// <summary>
    /// Builds DrawingsPart + ChartPart relationships for native Excel charts.
    /// </summary>
    internal static class ChartPartBuilder
    {
        private const uint CatAxId = 1;
        private const uint ValAxId = 2;
        private const uint SecondaryValAxId = 3;

        public static void AttachCharts(WorksheetPart worksheetPart, ExcelWorksheet model) =>
            DrawingPartBuilder.Attach(worksheetPart, model);

        internal static ChartSpace BuildChartSpace(ExcelChart chart, ExcelWorksheet sheet)
        {
            var chartSpace = new ChartSpace();
            chartSpace.AppendChild(new Date1904 { Val = false });
            chartSpace.AppendChild(new EditingLanguage { Val = "en-US" });
            chartSpace.AppendChild(new RoundedCorners { Val = false });

            var cChart = new C.Chart();
            if (!string.IsNullOrWhiteSpace(chart.Title))
            {
                cChart.AppendChild(BuildTitle(chart.Title!));
            }

            cChart.AppendChild(new AutoTitleDeleted { Val = string.IsNullOrWhiteSpace(chart.Title) });

            var plotArea = new PlotArea();
            plotArea.AppendChild(new Layout());

            bool needsAxes = RequiresAxes(chart);
            bool needsSecondary = chart.Series.Any(s => s.UseSecondaryAxis);

            AppendPlotCharts(plotArea, chart, sheet);

            if (needsAxes)
            {
                if (IsScatterFamily(chart))
                {
                    plotArea.AppendChild(BuildValueAxis(CatAxId, ValAxId, chart.CategoryAxis, AxisPositionValues.Bottom, includeMajorGrid: false));
                    plotArea.AppendChild(BuildValueAxis(ValAxId, CatAxId, chart.ValueAxis, AxisPositionValues.Left));
                }
                else if (IsBubble(chart))
                {
                    plotArea.AppendChild(BuildValueAxis(CatAxId, ValAxId, chart.CategoryAxis, AxisPositionValues.Bottom, includeMajorGrid: false));
                    plotArea.AppendChild(BuildValueAxis(ValAxId, CatAxId, chart.ValueAxis, AxisPositionValues.Left));
                }
                else
                {
                    plotArea.AppendChild(BuildCategoryAxis(CatAxId, ValAxId, chart.CategoryAxis));
                    plotArea.AppendChild(BuildValueAxis(ValAxId, CatAxId, chart.ValueAxis, AxisPositionValues.Left));
                    if (needsSecondary)
                    {
                        plotArea.AppendChild(BuildValueAxis(
                            SecondaryValAxId,
                            CatAxId,
                            chart.SecondaryValueAxis,
                            AxisPositionValues.Right,
                            crossesMax: true));
                    }
                }
            }

            cChart.AppendChild(plotArea);

            if (chart.ShowLegend)
            {
                cChart.AppendChild(new Legend(
                    new LegendPosition { Val = ToLegendPosition(chart.LegendPosition) },
                    new Layout(),
                    new Overlay { Val = false }));
            }

            cChart.AppendChild(new PlotVisibleOnly { Val = true });
            cChart.AppendChild(new DisplayBlanksAs { Val = DisplayBlanksAsValues.Gap });
            cChart.AppendChild(new ShowDataLabelsOverMaximum { Val = false });

            chartSpace.AppendChild(cChart);
            return chartSpace;
        }

        private static void AppendPlotCharts(PlotArea plotArea, ExcelChart chart, ExcelWorksheet sheet)
        {
            if (chart.Type == ExcelChartType.Combo)
            {
                AppendComboPlots(plotArea, chart, sheet);
                return;
            }

            ExcelChartType type = chart.Type;
            IList<ExcelChartSeries> series = chart.Series.ToList();
            uint startIndex = 0;
            OpenXmlCompositeElement? plot = BuildTypedPlot(type, chart, sheet, series, ref startIndex, forcePrimaryOnly: false);
            if (plot != null)
            {
                plotArea.AppendChild(plot);
            }
        }

        private static void AppendComboPlots(PlotArea plotArea, ExcelChart chart, ExcelWorksheet sheet)
        {
            // Group by (effective type, secondary flag) so each OpenXml chart element gets correct axIds.
            var groups = chart.Series
                .Select((s, i) => new { Series = s, Index = (uint)i })
                .GroupBy(x => new
                {
                    Type = chart.EffectiveSeriesType(x.Series),
                    Secondary = x.Series.UseSecondaryAxis
                });

            foreach (var group in groups)
            {
                var list = group.Select(g => g.Series).ToList();
                uint idx = group.Min(g => g.Index);
                // Preserve original indices: build one series at a time with absolute indexes.
                OpenXmlCompositeElement? plot = BuildTypedPlot(
                    group.Key.Type,
                    chart,
                    sheet,
                    list,
                    ref idx,
                    forcePrimaryOnly: false,
                    absoluteIndexes: group.Select(g => g.Index).ToList(),
                    forceSecondary: group.Key.Secondary);
                if (plot != null)
                {
                    plotArea.AppendChild(plot);
                }
            }
        }

        private static OpenXmlCompositeElement? BuildTypedPlot(
            ExcelChartType type,
            ExcelChart chart,
            ExcelWorksheet sheet,
            IList<ExcelChartSeries> series,
            ref uint index,
            bool forcePrimaryOnly,
            IList<uint>? absoluteIndexes = null,
            bool? forceSecondary = null)
        {
            switch (type)
            {
                case ExcelChartType.ColumnClustered:
                case ExcelChartType.ColumnStacked:
                case ExcelChartType.ColumnStacked100:
                    return BuildBarChart(chart, sheet, series, ref index, true, ToBarGrouping(type), absoluteIndexes, forceSecondary);
                case ExcelChartType.BarClustered:
                case ExcelChartType.BarStacked:
                case ExcelChartType.BarStacked100:
                    return BuildBarChart(chart, sheet, series, ref index, false, ToBarGrouping(type), absoluteIndexes, forceSecondary);
                case ExcelChartType.Line:
                case ExcelChartType.LineMarker:
                case ExcelChartType.LineStacked:
                case ExcelChartType.LineStacked100:
                    return BuildLineChart(chart, sheet, series, ref index, ToLineGrouping(type), absoluteIndexes, forceSecondary);
                case ExcelChartType.Scatter:
                    return BuildScatterChart(chart, sheet, series, ref index, absoluteIndexes);
                case ExcelChartType.Area:
                case ExcelChartType.AreaStacked:
                case ExcelChartType.AreaStacked100:
                    return BuildAreaChart(chart, sheet, series, ref index, ToAreaGrouping(type), absoluteIndexes, forceSecondary);
                case ExcelChartType.Pie:
                    return BuildPieChart(chart, sheet, series, ref index, absoluteIndexes);
                case ExcelChartType.Doughnut:
                    return BuildDoughnutChart(chart, sheet, series, ref index, absoluteIndexes);
                case ExcelChartType.Bubble:
                    return BuildBubbleChart(chart, sheet, series, ref index, absoluteIndexes);
                case ExcelChartType.Radar:
                    return BuildRadarChart(chart, sheet, series, ref index, absoluteIndexes, forceSecondary);
                case ExcelChartType.BoxAndWhisker:
                    throw new NotSupportedException("BoxAndWhisker is written via ChartEx, not a classic plot element.");
                default:
                    throw new NotSupportedException($"Chart type '{type}' is not supported.");
            }
        }

        private static BarGroupingValues ToBarGrouping(ExcelChartType type)
        {
            switch (type)
            {
                case ExcelChartType.ColumnStacked:
                case ExcelChartType.BarStacked:
                    return BarGroupingValues.Stacked;
                case ExcelChartType.ColumnStacked100:
                case ExcelChartType.BarStacked100:
                    return BarGroupingValues.PercentStacked;
                default:
                    return BarGroupingValues.Clustered;
            }
        }

        private static GroupingValues ToLineGrouping(ExcelChartType type)
        {
            switch (type)
            {
                case ExcelChartType.LineStacked:
                    return GroupingValues.Stacked;
                case ExcelChartType.LineStacked100:
                    return GroupingValues.PercentStacked;
                default:
                    return GroupingValues.Standard;
            }
        }

        private static GroupingValues ToAreaGrouping(ExcelChartType type)
        {
            switch (type)
            {
                case ExcelChartType.AreaStacked:
                    return GroupingValues.Stacked;
                case ExcelChartType.AreaStacked100:
                    return GroupingValues.PercentStacked;
                default:
                    return GroupingValues.Standard;
            }
        }

        private static bool RequiresAxes(ExcelChart chart)
        {
            if (chart.Type == ExcelChartType.Pie || chart.Type == ExcelChartType.Doughnut)
            {
                return false;
            }

            if (chart.Type == ExcelChartType.Combo)
            {
                return chart.Series.Any(s =>
                {
                    ExcelChartType t = chart.EffectiveSeriesType(s);
                    return t != ExcelChartType.Pie && t != ExcelChartType.Doughnut;
                });
            }

            return true;
        }

        private static bool IsScatterFamily(ExcelChart chart) =>
            chart.Type == ExcelChartType.Scatter
            || (chart.Type == ExcelChartType.Combo && chart.Series.All(s => chart.EffectiveSeriesType(s) == ExcelChartType.Scatter));

        private static bool IsBubble(ExcelChart chart) =>
            chart.Type == ExcelChartType.Bubble
            || (chart.Type == ExcelChartType.Combo && chart.Series.All(s => chart.EffectiveSeriesType(s) == ExcelChartType.Bubble));

        private static BarChart BuildBarChart(
            ExcelChart chart,
            ExcelWorksheet sheet,
            IList<ExcelChartSeries> seriesList,
            ref uint index,
            bool column,
            BarGroupingValues grouping,
            IList<uint>? absoluteIndexes,
            bool? forceSecondary)
        {
            var barChart = new BarChart(
                new BarDirection { Val = column ? BarDirectionValues.Column : BarDirectionValues.Bar },
                new BarGrouping { Val = grouping },
                new VaryColors { Val = false });

            for (int i = 0; i < seriesList.Count; i++)
            {
                uint serIndex = absoluteIndexes != null ? absoluteIndexes[i] : index++;
                if (absoluteIndexes == null)
                {
                    // index already incremented
                }
                else
                {
                    index = Math.Max(index, serIndex + 1);
                }

                ExcelChartSeries series = seriesList[i];
                BarChartSeries ser = BuildBarSeries(serIndex, series, sheet, chart);
                barChart.AppendChild(ser);
            }

            if (absoluteIndexes == null)
            {
                // ok
            }

            bool secondary = forceSecondary ?? seriesList.Any(s => s.UseSecondaryAxis);
            barChart.AppendChild(new AxisId { Val = CatAxId });
            barChart.AppendChild(new AxisId { Val = secondary ? SecondaryValAxId : ValAxId });
            return barChart;
        }

        private static LineChart BuildLineChart(
            ExcelChart chart,
            ExcelWorksheet sheet,
            IList<ExcelChartSeries> seriesList,
            ref uint index,
            GroupingValues grouping,
            IList<uint>? absoluteIndexes,
            bool? forceSecondary)
        {
            var lineChart = new LineChart(
                new Grouping { Val = grouping },
                new VaryColors { Val = false });

            bool anyMarkers = false;
            for (int i = 0; i < seriesList.Count; i++)
            {
                uint serIndex = absoluteIndexes != null ? absoluteIndexes[i] : index++;
                if (absoluteIndexes != null)
                {
                    index = Math.Max(index, serIndex + 1);
                }

                ExcelChartSeries series = seriesList[i];
                if (series.ShowMarkers && series.MarkerStyle != ExcelMarkerStyle.None)
                {
                    anyMarkers = true;
                }

                lineChart.AppendChild(BuildLineSeries(serIndex, series, sheet, chart));
            }

            lineChart.AppendChild(new ShowMarker { Val = anyMarkers });
            lineChart.AppendChild(new Smooth { Val = false });
            bool secondary = forceSecondary ?? seriesList.Any(s => s.UseSecondaryAxis);
            lineChart.AppendChild(new AxisId { Val = CatAxId });
            lineChart.AppendChild(new AxisId { Val = secondary ? SecondaryValAxId : ValAxId });
            return lineChart;
        }

        private static AreaChart BuildAreaChart(
            ExcelChart chart,
            ExcelWorksheet sheet,
            IList<ExcelChartSeries> seriesList,
            ref uint index,
            GroupingValues grouping,
            IList<uint>? absoluteIndexes,
            bool? forceSecondary)
        {
            var areaChart = new AreaChart(
                new Grouping { Val = grouping },
                new VaryColors { Val = false });

            for (int i = 0; i < seriesList.Count; i++)
            {
                uint serIndex = absoluteIndexes != null ? absoluteIndexes[i] : index++;
                if (absoluteIndexes != null)
                {
                    index = Math.Max(index, serIndex + 1);
                }

                areaChart.AppendChild(BuildAreaSeries(serIndex, seriesList[i], sheet, chart));
            }

            bool secondary = forceSecondary ?? seriesList.Any(s => s.UseSecondaryAxis);
            areaChart.AppendChild(new AxisId { Val = CatAxId });
            areaChart.AppendChild(new AxisId { Val = secondary ? SecondaryValAxId : ValAxId });
            return areaChart;
        }

        private static PieChart BuildPieChart(
            ExcelChart chart,
            ExcelWorksheet sheet,
            IList<ExcelChartSeries> seriesList,
            ref uint index,
            IList<uint>? absoluteIndexes)
        {
            var pie = new PieChart(new VaryColors { Val = true });
            for (int i = 0; i < seriesList.Count; i++)
            {
                uint serIndex = absoluteIndexes != null ? absoluteIndexes[i] : index++;
                if (absoluteIndexes != null)
                {
                    index = Math.Max(index, serIndex + 1);
                }

                pie.AppendChild(BuildPieSeries(serIndex, seriesList[i], sheet, chart));
            }

            pie.AppendChild(BuildDataLabels(ResolveLabels(chart, seriesList.FirstOrDefault()), pieDefaults: true));
            return pie;
        }

        private static DoughnutChart BuildDoughnutChart(
            ExcelChart chart,
            ExcelWorksheet sheet,
            IList<ExcelChartSeries> seriesList,
            ref uint index,
            IList<uint>? absoluteIndexes)
        {
            var donut = new DoughnutChart(new VaryColors { Val = true });
            for (int i = 0; i < seriesList.Count; i++)
            {
                uint serIndex = absoluteIndexes != null ? absoluteIndexes[i] : index++;
                if (absoluteIndexes != null)
                {
                    index = Math.Max(index, serIndex + 1);
                }

                donut.AppendChild(BuildPieSeries(serIndex, seriesList[i], sheet, chart));
            }

            donut.AppendChild(BuildDataLabels(ResolveLabels(chart, seriesList.FirstOrDefault()), pieDefaults: true));
            int hole = chart.DoughnutHoleSize;
            if (hole < 0) hole = 0;
            if (hole > 90) hole = 90;
            donut.AppendChild(new HoleSize { Val = (byte)hole });
            return donut;
        }

        private static ScatterChart BuildScatterChart(
            ExcelChart chart,
            ExcelWorksheet sheet,
            IList<ExcelChartSeries> seriesList,
            ref uint index,
            IList<uint>? absoluteIndexes)
        {
            var scatterChart = new ScatterChart(
                new ScatterStyle { Val = ScatterStyleValues.LineMarker },
                new VaryColors { Val = false });

            for (int i = 0; i < seriesList.Count; i++)
            {
                uint serIndex = absoluteIndexes != null ? absoluteIndexes[i] : index++;
                if (absoluteIndexes != null)
                {
                    index = Math.Max(index, serIndex + 1);
                }

                scatterChart.AppendChild(BuildScatterSeries(serIndex, seriesList[i], sheet, chart));
            }

            scatterChart.AppendChild(new AxisId { Val = CatAxId });
            scatterChart.AppendChild(new AxisId { Val = ValAxId });
            return scatterChart;
        }

        private static BubbleChart BuildBubbleChart(
            ExcelChart chart,
            ExcelWorksheet sheet,
            IList<ExcelChartSeries> seriesList,
            ref uint index,
            IList<uint>? absoluteIndexes)
        {
            var bubble = new BubbleChart(new VaryColors { Val = false });
            for (int i = 0; i < seriesList.Count; i++)
            {
                uint serIndex = absoluteIndexes != null ? absoluteIndexes[i] : index++;
                if (absoluteIndexes != null)
                {
                    index = Math.Max(index, serIndex + 1);
                }

                bubble.AppendChild(BuildBubbleSeries(serIndex, seriesList[i], sheet, chart));
            }

            bubble.AppendChild(new BubbleScale { Val = 100 });
            bubble.AppendChild(new ShowNegativeBubbles { Val = true });
            bubble.AppendChild(new AxisId { Val = CatAxId });
            bubble.AppendChild(new AxisId { Val = ValAxId });
            return bubble;
        }

        private static RadarChart BuildRadarChart(
            ExcelChart chart,
            ExcelWorksheet sheet,
            IList<ExcelChartSeries> seriesList,
            ref uint index,
            IList<uint>? absoluteIndexes,
            bool? forceSecondary)
        {
            var radar = new RadarChart(
                new RadarStyle { Val = RadarStyleValues.Marker },
                new VaryColors { Val = false });

            for (int i = 0; i < seriesList.Count; i++)
            {
                uint serIndex = absoluteIndexes != null ? absoluteIndexes[i] : index++;
                if (absoluteIndexes != null)
                {
                    index = Math.Max(index, serIndex + 1);
                }

                radar.AppendChild(BuildRadarSeries(serIndex, seriesList[i], sheet, chart));
            }

            bool secondary = forceSecondary ?? seriesList.Any(s => s.UseSecondaryAxis);
            radar.AppendChild(new AxisId { Val = CatAxId });
            radar.AppendChild(new AxisId { Val = secondary ? SecondaryValAxId : ValAxId });
            return radar;
        }

        private static BarChartSeries BuildBarSeries(uint index, ExcelChartSeries series, ExcelWorksheet sheet, ExcelChart chart)
        {
            // CT_BarSer: idx, order, tx?, ..., dLbls?, ..., cat?, val?
            var ser = new BarChartSeries();
            ser.AppendChild(new Index { Val = index });
            ser.AppendChild(new Order { Val = index });
            AppendSeriesText(ser, series, sheet);
            AppendSeriesColor(ser, series, chart.EffectiveSeriesType(series));
            AppendSeriesDataLabels(ser, chart, series);
            if (!string.IsNullOrWhiteSpace(series.CategoriesRange))
            {
                ser.AppendChild(new CategoryAxisData(BuildStringReference(sheet, series.CategoriesRange!)));
            }

            ser.AppendChild(new Values(BuildNumberReference(sheet, series.ValuesRange!)));
            return ser;
        }

        private static LineChartSeries BuildLineSeries(uint index, ExcelChartSeries series, ExcelWorksheet sheet, ExcelChart chart)
        {
            // CT_LineSer: idx, order, tx?, marker?, dLbls?, cat?, val?, smooth?
            var ser = new LineChartSeries();
            ser.AppendChild(new Index { Val = index });
            ser.AppendChild(new Order { Val = index });
            AppendSeriesText(ser, series, sheet);
            AppendSeriesColor(ser, series, chart.EffectiveSeriesType(series));
            AppendMarker(ser, series);
            AppendSeriesDataLabels(ser, chart, series);
            if (!string.IsNullOrWhiteSpace(series.CategoriesRange))
            {
                ser.AppendChild(new CategoryAxisData(BuildStringReference(sheet, series.CategoriesRange!)));
            }

            ser.AppendChild(new Values(BuildNumberReference(sheet, series.ValuesRange!)));
            return ser;
        }

        private static AreaChartSeries BuildAreaSeries(uint index, ExcelChartSeries series, ExcelWorksheet sheet, ExcelChart chart)
        {
            var ser = new AreaChartSeries();
            ser.AppendChild(new Index { Val = index });
            ser.AppendChild(new Order { Val = index });
            AppendSeriesText(ser, series, sheet);
            AppendSeriesColor(ser, series, chart.EffectiveSeriesType(series));
            AppendSeriesDataLabels(ser, chart, series);
            if (!string.IsNullOrWhiteSpace(series.CategoriesRange))
            {
                ser.AppendChild(new CategoryAxisData(BuildStringReference(sheet, series.CategoriesRange!)));
            }

            ser.AppendChild(new Values(BuildNumberReference(sheet, series.ValuesRange!)));
            return ser;
        }

        private static PieChartSeries BuildPieSeries(uint index, ExcelChartSeries series, ExcelWorksheet sheet, ExcelChart chart)
        {
            var ser = new PieChartSeries();
            ser.AppendChild(new Index { Val = index });
            ser.AppendChild(new Order { Val = index });
            AppendSeriesText(ser, series, sheet);
            AppendSeriesColor(ser, series, chart.EffectiveSeriesType(series));
            AppendSeriesDataLabels(ser, chart, series, pieDefaults: true);
            if (!string.IsNullOrWhiteSpace(series.CategoriesRange))
            {
                ser.AppendChild(new CategoryAxisData(BuildStringReference(sheet, series.CategoriesRange!)));
            }

            ser.AppendChild(new Values(BuildNumberReference(sheet, series.ValuesRange!)));
            return ser;
        }

        private static ScatterChartSeries BuildScatterSeries(uint index, ExcelChartSeries series, ExcelWorksheet sheet, ExcelChart chart)
        {
            var ser = new ScatterChartSeries();
            ser.AppendChild(new Index { Val = index });
            ser.AppendChild(new Order { Val = index });
            AppendSeriesText(ser, series, sheet);
            AppendSeriesColor(ser, series, chart.EffectiveSeriesType(series));
            AppendMarker(ser, series);
            AppendSeriesDataLabels(ser, chart, series);
            string xRange = series.CategoriesRange ?? series.ValuesRange!;
            ser.AppendChild(new XValues(BuildNumberReference(sheet, xRange)));
            ser.AppendChild(new YValues(BuildNumberReference(sheet, series.ValuesRange!)));
            ser.AppendChild(new Smooth { Val = false });
            return ser;
        }

        private static BubbleChartSeries BuildBubbleSeries(uint index, ExcelChartSeries series, ExcelWorksheet sheet, ExcelChart chart)
        {
            var ser = new BubbleChartSeries();
            ser.AppendChild(new Index { Val = index });
            ser.AppendChild(new Order { Val = index });
            AppendSeriesText(ser, series, sheet);
            AppendSeriesColor(ser, series, chart.EffectiveSeriesType(series));
            AppendSeriesDataLabels(ser, chart, series);
            string xRange = series.CategoriesRange ?? series.ValuesRange!;
            string sizeRange = series.BubbleSizeRange ?? series.ValuesRange!;
            ser.AppendChild(new XValues(BuildNumberReference(sheet, xRange)));
            ser.AppendChild(new YValues(BuildNumberReference(sheet, series.ValuesRange!)));
            ser.AppendChild(new BubbleSize(BuildNumberReference(sheet, sizeRange)));
            ser.AppendChild(new Bubble3D { Val = false });
            return ser;
        }

        private static RadarChartSeries BuildRadarSeries(uint index, ExcelChartSeries series, ExcelWorksheet sheet, ExcelChart chart)
        {
            var ser = new RadarChartSeries();
            ser.AppendChild(new Index { Val = index });
            ser.AppendChild(new Order { Val = index });
            AppendSeriesText(ser, series, sheet);
            AppendSeriesColor(ser, series, chart.EffectiveSeriesType(series));
            AppendMarker(ser, series);
            AppendSeriesDataLabels(ser, chart, series);
            if (!string.IsNullOrWhiteSpace(series.CategoriesRange))
            {
                ser.AppendChild(new CategoryAxisData(BuildStringReference(sheet, series.CategoriesRange!)));
            }

            ser.AppendChild(new Values(BuildNumberReference(sheet, series.ValuesRange!)));
            return ser;
        }

        private static void AppendMarker(OpenXmlCompositeElement seriesElement, ExcelChartSeries series)
        {
            if (!series.ShowMarkers || series.MarkerStyle == ExcelMarkerStyle.None)
            {
                seriesElement.AppendChild(new Marker(new Symbol { Val = MarkerStyleValues.None }));
                return;
            }

            seriesElement.AppendChild(new Marker(
                new Symbol { Val = ToMarkerStyle(series.MarkerStyle) },
                new Size { Val = 7 }));
        }

        private static MarkerStyleValues ToMarkerStyle(ExcelMarkerStyle style)
        {
            switch (style)
            {
                case ExcelMarkerStyle.Dash: return MarkerStyleValues.Dash;
                case ExcelMarkerStyle.Diamond: return MarkerStyleValues.Diamond;
                case ExcelMarkerStyle.Dot: return MarkerStyleValues.Dot;
                case ExcelMarkerStyle.Plus: return MarkerStyleValues.Plus;
                case ExcelMarkerStyle.Square: return MarkerStyleValues.Square;
                case ExcelMarkerStyle.Star: return MarkerStyleValues.Star;
                case ExcelMarkerStyle.Triangle: return MarkerStyleValues.Triangle;
                case ExcelMarkerStyle.X: return MarkerStyleValues.X;
                case ExcelMarkerStyle.Circle:
                case ExcelMarkerStyle.Automatic:
                default:
                    return MarkerStyleValues.Circle;
            }
        }

        private static ExcelDataLabelOptions ResolveLabels(ExcelChart chart, ExcelChartSeries? series)
        {
            if (series != null && series.DataLabels.Show)
            {
                return series.DataLabels;
            }

            return chart.DataLabels;
        }

        private static void AppendSeriesDataLabels(
            OpenXmlCompositeElement seriesElement,
            ExcelChart chart,
            ExcelChartSeries series,
            bool pieDefaults = false)
        {
            ExcelDataLabelOptions opts = ResolveLabels(chart, series);
            if (!opts.Show && !pieDefaults)
            {
                return;
            }

            if (!opts.Show && pieDefaults)
            {
                // Pie chart-level dLbls is enough; skip empty series labels.
                return;
            }

            seriesElement.AppendChild(BuildDataLabels(opts, pieDefaults));
        }

        private static DataLabels BuildDataLabels(ExcelDataLabelOptions opts, bool pieDefaults)
        {
            bool showValue = opts.Show && opts.ShowValue;
            bool showCat = opts.Show && opts.ShowCategoryName;
            bool showSer = opts.Show && opts.ShowSeriesName;
            bool showPct = opts.Show && (opts.ShowPercent || (pieDefaults && opts.Show && !opts.ShowValue && !opts.ShowCategoryName && !opts.ShowSeriesName));

            if (pieDefaults && opts.Show && !opts.ShowValue && !opts.ShowCategoryName && !opts.ShowSeriesName && !opts.ShowPercent)
            {
                showPct = true;
                showCat = true;
            }

            var labels = new DataLabels();
            if (!string.IsNullOrEmpty(opts.NumberFormat))
            {
                labels.AppendChild(new NumberingFormat
                {
                    FormatCode = opts.NumberFormat,
                    SourceLinked = false
                });
            }

            labels.AppendChild(new ShowLegendKey { Val = false });
            labels.AppendChild(new ShowValue { Val = showValue });
            labels.AppendChild(new ShowCategoryName { Val = showCat });
            labels.AppendChild(new ShowSeriesName { Val = showSer });
            labels.AppendChild(new ShowPercent { Val = showPct });
            labels.AppendChild(new ShowBubbleSize { Val = false });
            if (!string.IsNullOrEmpty(opts.Separator))
            {
                labels.AppendChild(new Separator(opts.Separator!));
            }

            labels.AppendChild(new ShowLeaderLines { Val = pieDefaults });
            return labels;
        }

        private static void AppendSeriesColor(
            OpenXmlCompositeElement seriesElement,
            ExcelChartSeries series,
            ExcelChartType type)
        {
            if (series.Color is null)
            {
                return;
            }

            string rgb = series.Color.ToRgbHex();
            var spPr = new ChartShapeProperties();
            if (UsesOutlineColor(type))
            {
                spPr.AppendChild(new A.Outline(
                    new A.SolidFill(new A.RgbColorModelHex { Val = rgb }))
                {
                    Width = 19050
                });
            }
            else
            {
                spPr.AppendChild(new A.SolidFill(new A.RgbColorModelHex { Val = rgb }));
            }

            seriesElement.AppendChild(spPr);
        }

        private static bool UsesOutlineColor(ExcelChartType type)
        {
            switch (type)
            {
                case ExcelChartType.Line:
                case ExcelChartType.LineMarker:
                case ExcelChartType.LineStacked:
                case ExcelChartType.LineStacked100:
                case ExcelChartType.Scatter:
                case ExcelChartType.Radar:
                    return true;
                default:
                    return false;
            }
        }

        private static void AppendSeriesText(OpenXmlCompositeElement seriesElement, ExcelChartSeries series, ExcelWorksheet sheet)
        {
            if (!string.IsNullOrWhiteSpace(series.NameRange))
            {
                string formula = ChartFormulaHelper.ToSheetLocalFormula(sheet.Name, series.NameRange!);
                List<object?> values = ChartFormulaHelper.ReadLinearValues(sheet, series.NameRange!);
                string text = ChartFormulaHelper.ToCacheString(values.Count > 0 ? values[0] : null) ?? series.Name ?? "Series";

                seriesElement.AppendChild(new SeriesText(
                    new StringReference(
                        new Formula(formula),
                        BuildStringCache(new List<string?> { text }))));
                return;
            }

            if (!string.IsNullOrWhiteSpace(series.Name))
            {
                seriesElement.AppendChild(new SeriesText(new NumericValue(series.Name!)));
            }
        }

        private static StringReference BuildStringReference(ExcelWorksheet sheet, string a1Range)
        {
            string formula = ChartFormulaHelper.ToSheetLocalFormula(sheet.Name, a1Range);
            List<object?> raw = ChartFormulaHelper.ReadLinearValues(sheet, a1Range);
            var strings = new List<string?>();
            foreach (object? item in raw)
            {
                strings.Add(ChartFormulaHelper.ToCacheString(item) ?? string.Empty);
            }

            return new StringReference(new Formula(formula), BuildStringCache(strings));
        }

        private static NumberReference BuildNumberReference(ExcelWorksheet sheet, string a1Range)
        {
            string formula = ChartFormulaHelper.ToSheetLocalFormula(sheet.Name, a1Range);
            List<object?> raw = ChartFormulaHelper.ReadLinearValues(sheet, a1Range);
            var numbers = new List<double?>();
            foreach (object? item in raw)
            {
                numbers.Add(ChartFormulaHelper.ToCacheNumber(item));
            }

            return new NumberReference(new Formula(formula), BuildNumberingCache(numbers));
        }

        private static StringCache BuildStringCache(IReadOnlyList<string?> values)
        {
            var cache = new StringCache();
            cache.AppendChild(new PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++)
            {
                cache.AppendChild(new StringPoint
                {
                    Index = (uint)i,
                    NumericValue = new NumericValue(values[i] ?? string.Empty)
                });
            }

            return cache;
        }

        private static NumberingCache BuildNumberingCache(IReadOnlyList<double?> values)
        {
            var cache = new NumberingCache();
            cache.AppendChild(new FormatCode("General"));
            cache.AppendChild(new PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++)
            {
                if (!values[i].HasValue)
                {
                    continue;
                }

                cache.AppendChild(new NumericPoint
                {
                    Index = (uint)i,
                    NumericValue = new NumericValue(values[i]!.Value.ToString("G17", CultureInfo.InvariantCulture))
                });
            }

            return cache;
        }

        private static Title BuildTitle(string text)
        {
            return new Title(
                new ChartText(
                    new RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.ParagraphProperties(new A.DefaultRunProperties()),
                            new A.Run(
                                new A.RunProperties { Language = "en-US" },
                                new A.Text(text)),
                            new A.EndParagraphRunProperties { Language = "en-US" }))),
                new Overlay { Val = false });
        }

        private static Title BuildAxisTitle(string text)
        {
            return BuildTitle(text);
        }

        private static CategoryAxis BuildCategoryAxis(uint axId, uint crossAxId, ExcelChartAxis options)
        {
            // CT_CatAx: axId, scaling, delete, axPos, majorGridlines?, minorGridlines?, title?, ...
            var axis = new CategoryAxis(
                new AxisId { Val = axId },
                new Scaling(new Orientation { Val = OrientationValues.MinMax }),
                new Delete { Val = false },
                new AxisPosition { Val = AxisPositionValues.Bottom });

            if (options.ShowMajorGridlines)
            {
                axis.AppendChild(new MajorGridlines());
            }

            if (options.ShowMinorGridlines)
            {
                axis.AppendChild(new MinorGridlines());
            }

            if (!string.IsNullOrWhiteSpace(options.Title))
            {
                axis.AppendChild(BuildAxisTitle(options.Title!));
            }

            axis.AppendChild(new MajorTickMark { Val = TickMarkValues.Outside });
            axis.AppendChild(new MinorTickMark { Val = TickMarkValues.None });
            axis.AppendChild(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
            axis.AppendChild(new CrossingAxis { Val = crossAxId });
            axis.AppendChild(new Crosses { Val = CrossesValues.AutoZero });
            axis.AppendChild(new AutoLabeled { Val = true });
            axis.AppendChild(new LabelAlignment { Val = LabelAlignmentValues.Center });
            axis.AppendChild(new LabelOffset { Val = 100 });
            axis.AppendChild(new NoMultiLevelLabels { Val = false });
            return axis;
        }

        private static ValueAxis BuildValueAxis(
            uint axId,
            uint crossAxId,
            ExcelChartAxis options,
            AxisPositionValues position,
            bool includeMajorGrid = true,
            bool crossesMax = false)
        {
            // CT_Scaling: orientation, max, min (max before min).
            var scaling = new Scaling(new Orientation { Val = OrientationValues.MinMax });
            if (options.Maximum.HasValue)
            {
                scaling.AppendChild(new MaxAxisValue { Val = options.Maximum.Value });
            }

            if (options.Minimum.HasValue)
            {
                scaling.AppendChild(new MinAxisValue { Val = options.Minimum.Value });
            }

            // CT_ValAx order: axId, scaling, delete, axPos, majorGridlines?, minorGridlines?, title?, ...
            // then tick marks, tickLblPos, crossAx, crosses, crossBetween.
            var axis = new ValueAxis(
                new AxisId { Val = axId },
                scaling,
                new Delete { Val = false },
                new AxisPosition { Val = position });

            bool major = includeMajorGrid && options.ShowMajorGridlines;
            if (major)
            {
                axis.AppendChild(new MajorGridlines());
            }

            if (options.ShowMinorGridlines)
            {
                axis.AppendChild(new MinorGridlines());
            }

            if (!string.IsNullOrWhiteSpace(options.Title))
            {
                axis.AppendChild(BuildAxisTitle(options.Title!));
            }

            axis.AppendChild(new MajorTickMark { Val = TickMarkValues.Outside });
            axis.AppendChild(new MinorTickMark { Val = TickMarkValues.None });
            axis.AppendChild(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
            axis.AppendChild(new CrossingAxis { Val = crossAxId });
            axis.AppendChild(new Crosses { Val = crossesMax ? CrossesValues.Maximum : CrossesValues.AutoZero });
            axis.AppendChild(new CrossBetween { Val = CrossBetweenValues.Between });
            return axis;
        }

        private static LegendPositionValues ToLegendPosition(ExcelLegendPosition position)
        {
            switch (position)
            {
                case ExcelLegendPosition.Left: return LegendPositionValues.Left;
                case ExcelLegendPosition.Top: return LegendPositionValues.Top;
                case ExcelLegendPosition.Bottom: return LegendPositionValues.Bottom;
                case ExcelLegendPosition.TopRight: return LegendPositionValues.TopRight;
                default: return LegendPositionValues.Right;
            }
        }

        internal static Xdr.TwoCellAnchor CreateRangeAnchor(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            var anchor = new Xdr.TwoCellAnchor();
            anchor.AppendChild(new Xdr.FromMarker(
                new Xdr.ColumnId((fromColumn - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId((fromRow - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowOffset("0")));
            anchor.AppendChild(new Xdr.ToMarker(
                new Xdr.ColumnId((toColumn - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId((toRow - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowOffset("0")));
            return anchor;
        }

        internal static Xdr.TwoCellAnchor BuildTwoCellAnchor(ExcelChart chart, string chartRelId, uint shapeId)
        {
            var anchor = CreateRangeAnchor(chart.FromRow, chart.FromColumn, chart.ToRow, chart.ToColumn);

            var graphicFrame = new Xdr.GraphicFrame { Macro = string.Empty };
            graphicFrame.AppendChild(new Xdr.NonVisualGraphicFrameProperties(
                new Xdr.NonVisualDrawingProperties { Id = shapeId, Name = "Chart " + shapeId },
                new Xdr.NonVisualGraphicFrameDrawingProperties()));
            graphicFrame.AppendChild(new Xdr.Transform(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 0L, Cy = 0L }));
            graphicFrame.AppendChild(new A.Graphic(
                new A.GraphicData(
                    new ChartReference { Id = chartRelId })
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                }));

            anchor.AppendChild(graphicFrame);
            anchor.AppendChild(new Xdr.ClientData());
            return anchor;
        }
    }
}
