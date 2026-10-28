using System;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// One chart series bound to worksheet ranges (A1 notation, sheet-local).
    /// </summary>
    public sealed class ExcelChartSeries
    {
        internal ExcelChartSeries(ExcelChart chart)
        {
            Chart = chart ?? throw new ArgumentNullException(nameof(chart));
        }

        public ExcelChart Chart { get; }

        /// <summary>Optional literal series name (used when <see cref="NameRange"/> is null).</summary>
        public string? Name { get; set; }

        /// <summary>Optional single-cell or range for series title (e.g. B1).</summary>
        public string? NameRange { get; set; }

        /// <summary>
        /// Categories (most charts), X values (scatter), or unused for some pie cases.
        /// </summary>
        public string? CategoriesRange { get; set; }

        /// <summary>
        /// Values / Y values. Required for most chart types.
        /// </summary>
        public string? ValuesRange { get; set; }

        /// <summary>Bubble size range (bubble charts only).</summary>
        public string? BubbleSizeRange { get; set; }

        /// <summary>
        /// Per-series plot type for <see cref="ExcelChartType.Combo"/> charts.
        /// Ignored for non-combo charts (chart-level type wins).
        /// </summary>
        public ExcelChartType? SeriesChartType { get; set; }

        /// <summary>Plot against the secondary value axis.</summary>
        public bool UseSecondaryAxis { get; set; }

        /// <summary>Show markers (line / scatter). Default true.</summary>
        public bool ShowMarkers { get; set; } = true;

        public ExcelMarkerStyle MarkerStyle { get; set; } = ExcelMarkerStyle.Automatic;

        /// <summary>
        /// Series color (RRGGBB). Line-like series use outline fill; column/bar/area/pie/bubble use solid fill.
        /// </summary>
        public ExcelColor? Color { get; set; }

        /// <summary>Optional per-series data labels (overrides chart defaults when <see cref="ExcelDataLabelOptions.Show"/> is set).</summary>
        public ExcelDataLabelOptions DataLabels { get; } = new ExcelDataLabelOptions();

        public ExcelChartSeries SetName(string name)
        {
            Name = name;
            NameRange = null;
            return this;
        }

        public ExcelChartSeries SetNameRange(string a1)
        {
            NameRange = a1;
            return this;
        }

        public ExcelChartSeries SetCategories(string a1Range)
        {
            CategoriesRange = a1Range;
            return this;
        }

        public ExcelChartSeries SetValues(string a1Range)
        {
            ValuesRange = a1Range;
            return this;
        }

        public ExcelChartSeries SetBubbleSize(string a1Range)
        {
            BubbleSizeRange = a1Range;
            return this;
        }

        public ExcelChartSeries OnSecondaryAxis(bool value = true)
        {
            UseSecondaryAxis = value;
            return this;
        }

        public ExcelChartSeries AsChartType(ExcelChartType type)
        {
            SeriesChartType = type;
            return this;
        }

        public ExcelChartSeries SetColor(ExcelColor color)
        {
            Color = color ?? throw new ArgumentNullException(nameof(color));
            return this;
        }

        public ExcelChartSeries SetColor(string hex)
        {
            Color = ExcelColor.FromHex(hex);
            return this;
        }
    }
}
