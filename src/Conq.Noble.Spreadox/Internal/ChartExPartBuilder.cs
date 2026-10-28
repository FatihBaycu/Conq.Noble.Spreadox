using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace Conq.Noble.Spreadox.Internal
{
    /// <summary>
    /// Office 2016+ ChartEx (box-and-whisker). Excel computes quartiles from the bound raw range.
    /// </summary>
    internal static class ChartExPartBuilder
    {
        public static Cx.ChartSpace BuildBoxAndWhisker(ExcelChart chart, ExcelWorksheet sheet)
        {
            ExcelChartSeries series = chart.Series.Count > 0
                ? chart.Series[0]
                : throw new InvalidOperationException("BoxAndWhisker requires at least one series.");

            ResolveRanges(series, out ExcelRange catRange, out ExcelRange valRange);

            var chartData = new Cx.ChartData();
            var data = new Cx.Data { Id = 0U };
            data.AppendChild(BuildStringDimension(sheet, catRange));
            data.AppendChild(BuildNumericDimension(sheet, valRange));
            chartData.AppendChild(data);

            var cxChart = new Cx.Chart();
            if (!string.IsNullOrWhiteSpace(chart.Title))
            {
                cxChart.AppendChild(BuildTitle(chart.Title!));
            }

            var plotArea = new Cx.PlotArea();
            var region = new Cx.PlotAreaRegion();
            region.AppendChild(BuildSeries(chart, series, sheet));
            plotArea.AppendChild(region);
            plotArea.AppendChild(BuildCategoryAxis());
            plotArea.AppendChild(BuildValueAxis(chart.ValueAxis));
            cxChart.AppendChild(plotArea);

            if (chart.ShowLegend)
            {
                cxChart.AppendChild(new Cx.Legend
                {
                    Overlay = false,
                    Pos = ToSidePos(chart.LegendPosition)
                });
            }

            var space = new Cx.ChartSpace();
            space.AppendChild(chartData);
            space.AppendChild(cxChart);
            return space;
        }

        private static void ResolveRanges(ExcelChartSeries series, out ExcelRange catRange, out ExcelRange valRange)
        {
            if (string.IsNullOrWhiteSpace(series.ValuesRange))
            {
                throw new InvalidOperationException("BoxAndWhisker series requires a values range.");
            }

            ExcelRange values = ExcelRange.Parse(series.ValuesRange!);
            if (!string.IsNullOrWhiteSpace(series.CategoriesRange))
            {
                catRange = ExcelRange.Parse(series.CategoriesRange!);
                valRange = values;
                return;
            }

            // Header row + value block in one range (Analyst/BoQP convention).
            if (values.EndRow > values.StartRow)
            {
                catRange = new ExcelRange(values.StartRow, values.StartColumn, values.StartRow, values.EndColumn);
                valRange = new ExcelRange(values.StartRow + 1, values.StartColumn, values.EndRow, values.EndColumn);
                return;
            }

            catRange = values;
            valRange = values;
        }

        private static Cx.Series BuildSeries(ExcelChart chart, ExcelChartSeries series, ExcelWorksheet sheet)
        {
            var ser = new Cx.Series
            {
                LayoutId = Cx.SeriesLayout.BoxWhisker,
                UniqueId = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}"
            };

            string name = series.Name
                          ?? FirstHeader(sheet, series)
                          ?? "Series 1";
            ser.AppendChild(new Cx.Text(new Cx.TextData(new Cx.VXsdstring(name))));

            if (series.Color != null)
            {
                ser.AppendChild(new Cx.ShapeProperties(
                    new DocumentFormat.OpenXml.Drawing.SolidFill(
                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = series.Color.ToRgbHex() })));
            }

            ser.AppendChild(new Cx.DataId { Val = 0U });

            ExcelBoxAndWhiskerOptions opts = chart.BoxAndWhisker;
            var layout = new Cx.SeriesLayoutProperties();
            layout.AppendChild(new Cx.SeriesElementVisibilities
            {
                Outliers = opts.ShowOutlierPoints,
                Nonoutliers = opts.ShowInnerPoints,
                MeanMarker = opts.ShowMeanMarkers,
                MeanLine = opts.ShowMeanLine
            });
            layout.AppendChild(new Cx.Statistics { QuartileMethod = Cx.QuartileMethod.Inclusive });
            ser.AppendChild(layout);
            ser.AppendChild(new Cx.AxisId("0"));
            ser.AppendChild(new Cx.AxisId("1"));
            return ser;
        }

        private static string? FirstHeader(ExcelWorksheet sheet, ExcelChartSeries series)
        {
            if (!string.IsNullOrWhiteSpace(series.NameRange))
            {
                List<object?> values = ChartFormulaHelper.ReadLinearValues(sheet, series.NameRange!);
                return values.Count > 0 ? ChartFormulaHelper.ToCacheString(values[0]) : null;
            }

            return null;
        }

        private static Cx.StringDimension BuildStringDimension(ExcelWorksheet sheet, ExcelRange range)
        {
            var dim = new Cx.StringDimension { Type = Cx.StringDimensionType.Cat };
            dim.AppendChild(new Cx.Formula(ChartFormulaHelper.ToSheetLocalFormula(sheet.Name, range)));
            List<object?> values = ChartFormulaHelper.ReadLinearValues(sheet, range.A1);
            var level = new Cx.StringLevel { PtCount = (uint)values.Count };
            for (int i = 0; i < values.Count; i++)
            {
                level.AppendChild(new Cx.ChartStringValue
                {
                    Index = (uint)i,
                    Text = ChartFormulaHelper.ToCacheString(values[i]) ?? string.Empty
                });
            }

            dim.AppendChild(level);
            return dim;
        }

        private static Cx.NumericDimension BuildNumericDimension(ExcelWorksheet sheet, ExcelRange range)
        {
            var dim = new Cx.NumericDimension { Type = Cx.NumericDimensionType.Val };
            dim.AppendChild(new Cx.Formula(ChartFormulaHelper.ToSheetLocalFormula(sheet.Name, range)));

            int rows = range.EndRow - range.StartRow + 1;
            for (int c = range.StartColumn; c <= range.EndColumn; c++)
            {
                var level = new Cx.NumericLevel { PtCount = (uint)rows };
                for (int r = range.StartRow; r <= range.EndRow; r++)
                {
                    int idx = r - range.StartRow;
                    if (!sheet.TryGetCell(r, c, out ExcelCell? cell) || cell is null || cell.IsEmpty)
                    {
                        continue;
                    }

                    double? number = ChartFormulaHelper.ToCacheNumber(cell.Value);
                    if (!number.HasValue)
                    {
                        continue;
                    }

                    level.AppendChild(new Cx.NumericValue
                    {
                        Idx = (uint)idx,
                        Text = number.Value.ToString("G17", CultureInfo.InvariantCulture)
                    });
                }

                dim.AppendChild(level);
            }

            return dim;
        }

        private static Cx.Axis BuildCategoryAxis()
        {
            var axis = new Cx.Axis { Id = 0U };
            axis.AppendChild(new Cx.CategoryAxisScaling { GapWidth = "1" });
            axis.AppendChild(new Cx.TickLabels());
            return axis;
        }

        private static Cx.Axis BuildValueAxis(ExcelChartAxis options)
        {
            var axis = new Cx.Axis { Id = 1U };
            var scaling = new Cx.ValueAxisScaling();
            if (options.Minimum.HasValue)
            {
                scaling.Min = options.Minimum.Value.ToString(CultureInfo.InvariantCulture);
            }

            if (options.Maximum.HasValue)
            {
                scaling.Max = options.Maximum.Value.ToString(CultureInfo.InvariantCulture);
            }

            axis.AppendChild(scaling);
            if (options.ShowMajorGridlines)
            {
                axis.AppendChild(new Cx.MajorGridlinesGridlines());
            }

            if (!string.IsNullOrWhiteSpace(options.Title))
            {
                axis.AppendChild(new Cx.AxisTitle(
                    new Cx.Text(new Cx.TextData(new Cx.VXsdstring(options.Title!)))));
            }

            axis.AppendChild(new Cx.TickLabels());
            return axis;
        }

        private static Cx.ChartTitle BuildTitle(string text)
        {
            return new Cx.ChartTitle(
                new Cx.Text(new Cx.TextData(new Cx.VXsdstring(text))))
            {
                Overlay = false
            };
        }

        private static Cx.SidePos ToSidePos(ExcelLegendPosition position)
        {
            switch (position)
            {
                case ExcelLegendPosition.Left: return Cx.SidePos.L;
                case ExcelLegendPosition.Top: return Cx.SidePos.T;
                case ExcelLegendPosition.Bottom: return Cx.SidePos.B;
                default: return Cx.SidePos.R;
            }
        }
    }
}
