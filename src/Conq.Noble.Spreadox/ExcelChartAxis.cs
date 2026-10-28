namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Axis options for category / value / secondary value axes.
    /// </summary>
    public sealed class ExcelChartAxis
    {
        public string? Title { get; set; }

        public double? Minimum { get; set; }

        public double? Maximum { get; set; }

        public bool ShowMajorGridlines { get; set; } = true;

        public bool ShowMinorGridlines { get; set; }
    }

    /// <summary>
    /// Data label visibility flags (series or chart default).
    /// </summary>
    public sealed class ExcelDataLabelOptions
    {
        public bool Show { get; set; }

        public bool ShowValue { get; set; } = true;

        public bool ShowCategoryName { get; set; }

        public bool ShowSeriesName { get; set; }

        public bool ShowPercent { get; set; }

        /// <summary>Label separator when multiple parts are shown (e.g. <c>": "</c> for category + value).</summary>
        public string? Separator { get; set; }

        /// <summary>
        /// Excel number format for label values (e.g. <c>0.0"%"</c> for already-scaled 0–100 numbers).
        /// Do not use <c>0.0%</c> unless values are fractions.
        /// </summary>
        public string? NumberFormat { get; set; }

        public ExcelDataLabelOptions Clone()
        {
            return new ExcelDataLabelOptions
            {
                Show = Show,
                ShowValue = ShowValue,
                ShowCategoryName = ShowCategoryName,
                ShowSeriesName = ShowSeriesName,
                ShowPercent = ShowPercent,
                Separator = Separator,
                NumberFormat = NumberFormat
            };
        }
    }

    /// <summary>Box-and-whisker ChartEx options (Excel computes quartiles from the bound raw range).</summary>
    public sealed class ExcelBoxAndWhiskerOptions
    {
        public bool ShowOutlierPoints { get; set; } = true;

        public bool ShowMeanMarkers { get; set; } = true;

        public bool ShowMeanLine { get; set; }

        public bool ShowInnerPoints { get; set; } = true;
    }
}
