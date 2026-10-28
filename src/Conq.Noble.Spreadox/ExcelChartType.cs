namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Native OOXML chart types (DrawingML ChartPart — not images).
    /// </summary>
    public enum ExcelChartType
    {
        ColumnClustered,
        ColumnStacked,
        ColumnStacked100,
        BarClustered,
        BarStacked,
        BarStacked100,
        Line,
        LineStacked,
        LineStacked100,
        Scatter,
        Area,
        AreaStacked,
        AreaStacked100,
        Pie,
        Doughnut,
        Bubble,
        Radar,
        /// <summary>Multiple series types in one plotArea (set per-series <see cref="ExcelChartSeries.SeriesChartType"/>).</summary>
        Combo,
        /// <summary>Line chart with markers (OOXML <c>c:lineChart</c> + <c>c:marker</c>).</summary>
        LineMarker,
        /// <summary>Office 2016+ ChartEx box-and-whisker (Excel computes quartiles from raw values).</summary>
        BoxAndWhisker
    }

    public enum ExcelLegendPosition
    {
        Right,
        Left,
        Top,
        Bottom,
        TopRight
    }

    public enum ExcelMarkerStyle
    {
        Automatic,
        None,
        Circle,
        Dash,
        Diamond,
        Dot,
        Plus,
        Square,
        Star,
        Triangle,
        X
    }
}
