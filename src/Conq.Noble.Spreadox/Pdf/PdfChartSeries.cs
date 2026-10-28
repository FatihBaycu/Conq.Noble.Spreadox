using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Conq.Noble.Spreadox
{
    /// <summary>In-memory series for PDF charts (not worksheet ranges).</summary>
    public sealed class PdfChartSeries
    {
        internal PdfChartSeries(PdfChart chart, string? name, IReadOnlyList<string> categories, IReadOnlyList<double> values)
        {
            Chart = chart;
            Name = name;
            Categories = categories;
            Values = values;
        }

        public PdfChart Chart { get; }

        public string? Name { get; set; }

        public IReadOnlyList<string> Categories { get; }

        public IReadOnlyList<double> Values { get; }

        public ExcelChartType? SeriesChartType { get; set; }

        public bool UseSecondaryAxis { get; set; }

        public bool ShowMarkers { get; set; } = true;

        public ExcelColor? Color { get; set; }

        public PdfChartSeries AsChartType(ExcelChartType type)
        {
            SeriesChartType = type;
            return this;
        }

        public PdfChartSeries OnSecondaryAxis(bool value = true)
        {
            UseSecondaryAxis = value;
            return this;
        }

        public PdfChartSeries SetColor(ExcelColor color)
        {
            Color = color ?? throw new ArgumentNullException(nameof(color));
            return this;
        }

        public PdfChartSeries SetColor(string hex)
        {
            Color = ExcelColor.FromHex(hex);
            return this;
        }

        internal static IReadOnlyList<string> CopyCategories(IReadOnlyList<string>? categories, int valueCount)
        {
            var copy = new string[valueCount];
            for (int i = 0; i < valueCount; i++)
            {
                if (categories != null && i < categories.Count && categories[i] != null)
                {
                    copy[i] = categories[i];
                }
                else
                {
                    copy[i] = (i + 1).ToString();
                }
            }

            return new ReadOnlyCollection<string>(copy);
        }

        internal static IReadOnlyList<double> CopyValues(IReadOnlyList<double> values, int maxPoints)
        {
            if (values is null)
            {
                throw new ArgumentNullException(nameof(values));
            }

            int count = Math.Min(values.Count, maxPoints);
            var copy = new double[count];
            for (int i = 0; i < count; i++)
            {
                copy[i] = values[i];
            }

            return new ReadOnlyCollection<double>(copy);
        }
    }
}
