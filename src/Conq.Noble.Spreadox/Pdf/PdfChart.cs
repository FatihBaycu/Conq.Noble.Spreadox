using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// PDF chart drawn as vectors. Series are in-memory points, not Excel cell ranges.
    /// </summary>
    public sealed class PdfChart
    {
        private readonly List<PdfChartSeries> _series = new List<PdfChartSeries>();
        private readonly List<PdfConstantLine> _constantLines = new List<PdfConstantLine>();

        internal PdfChart()
        {
            Type = ExcelChartType.LineMarker;
            CategoryAxis.LabelRotation = -35;
            SecondaryValueAxis.Minimum = 0;
            SecondaryValueAxis.Maximum = 100;
            SecondaryValueAxis.Title = "%";
            SecondaryValueAxis.ShowMajorGridlines = false;
        }

        public ExcelChartType Type { get; set; }

        public bool ShowLegend { get; set; } = true;

        public ExcelLegendPosition LegendPosition { get; set; } = ExcelLegendPosition.Bottom;

        public PdfChartAxis CategoryAxis { get; } = new PdfChartAxis { ShowMajorGridlines = false, LabelRotation = -35 };

        public PdfChartAxis ValueAxis { get; } = new PdfChartAxis();

        public PdfChartAxis SecondaryValueAxis { get; } = new PdfChartAxis();

        /// <summary>Points kept per series (default 5000).</summary>
        public int MaxPoints { get; set; } = 5000;

        public IReadOnlyList<PdfChartSeries> Series => new ReadOnlyCollection<PdfChartSeries>(_series);

        public IReadOnlyList<PdfConstantLine> ConstantLines => new ReadOnlyCollection<PdfConstantLine>(_constantLines);

        public PdfChartSeries AddSeries(string? name, IReadOnlyList<string>? categories, IReadOnlyList<double> values)
        {
            if (MaxPoints < 1)
            {
                throw new InvalidOperationException("MaxPoints must be at least 1.");
            }

            IReadOnlyList<double> copiedValues = PdfChartSeries.CopyValues(values, MaxPoints);
            IReadOnlyList<string> copiedCategories = PdfChartSeries.CopyCategories(categories, copiedValues.Count);
            var series = new PdfChartSeries(this, name, copiedCategories, copiedValues);
            _series.Add(series);
            return series;
        }

        public PdfChartSeries AddSeries(string? name, IReadOnlyList<double> values)
        {
            return AddSeries(name, categories: null, values);
        }

        public PdfConstantLine AddConstantLine(string name, double value, ExcelColor color)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("Name is required.", nameof(name));
            }

            if (color is null)
            {
                throw new ArgumentNullException(nameof(color));
            }

            var line = new PdfConstantLine(name, value, color);
            _constantLines.Add(line);
            return line;
        }
    }
}
