using System.ComponentModel.DataAnnotations;
using Conq.Noble.Spreadox;

namespace ReportBuilderMvc.Models;

/// <summary>
/// Represents a single chart data series configuration.
/// </summary>
public sealed class ChartDataItem
{
    public string Title { get; set; } = string.Empty;

    public ExcelChartType ChartType { get; set; } = ExcelChartType.ColumnClustered;

    /// <summary>Comma-separated category labels.</summary>
    public string Categories { get; set; } = string.Empty;

    /// <summary>Comma-separated numeric values.</summary>
    public string Values { get; set; } = string.Empty;

    /// <summary>Series color as hex (e.g. #005293).</summary>
    public string Color { get; set; } = "#005293";
}

/// <summary>
/// Header configuration for PDF reports.
/// </summary>
public sealed class HeaderConfig
{
    public string? Title { get; set; }
    public string? Text { get; set; }
    public string? FontFamily { get; set; }
    public double FontSize { get; set; } = 12;
    public bool Bold { get; set; } = true;
    public string TitleColor { get; set; } = "#000000";
}

/// <summary>
/// Footer configuration for PDF reports.
/// </summary>
public sealed class FooterConfig
{
    public string? Text { get; set; }
    public string PageNumberFormat { get; set; } = "Page {0} / {1}";
    public string? FontFamily { get; set; }
}

/// <summary>
/// Cover page configuration for PDF reports.
/// </summary>
public sealed class CoverConfig
{
    public string? Company { get; set; }
    public string? Title { get; set; }
    public string? Provider { get; set; }
    public bool IncludeDate { get; set; } = true;
}

/// <summary>
/// Output format.
/// </summary>
public enum OutputFormat
{
    Excel,
    Pdf,
    Both
}

/// <summary>
/// Chart layout mode for PDF reports.
/// </summary>
public enum ChartLayout
{
    /// <summary>Each chart on its own page.</summary>
    SeparatePages,

    /// <summary>All charts on a single page (Excel sheet).</summary>
    SinglePage
}

/// <summary>
/// Main report configuration ViewModel.
/// </summary>
public sealed class ReportConfigViewModel
{
    [Display(Name = "Report Title")]
    public string ReportTitle { get; set; } = "Spreadox Sample Report";

    [Display(Name = "Output Format")]
    public OutputFormat Format { get; set; } = OutputFormat.Pdf;

    [Display(Name = "Chart Layout")]
    public ChartLayout Layout { get; set; } = ChartLayout.SeparatePages;

    [Display(Name = "Number of Data Series")]
    [Range(1, 20)]
    public int DataSeriesCount { get; set; } = 5;

    /// <summary>The chart data items (populated after generation).</summary>
    public List<ChartDataItem> Items { get; set; } = new();

    public HeaderConfig Header { get; set; } = new();

    public FooterConfig Footer { get; set; } = new();

    public CoverConfig Cover { get; set; } = new();

    /// <summary>Available chart types for the dropdown.</summary>
    public static IReadOnlyList<ExcelChartType> AvailableChartTypes { get; } = new[]
    {
        ExcelChartType.ColumnClustered,
        ExcelChartType.ColumnStacked,
        ExcelChartType.BarClustered,
        ExcelChartType.BarStacked,
        ExcelChartType.Line,
        ExcelChartType.LineMarker,
        ExcelChartType.LineStacked,
        ExcelChartType.Area,
        ExcelChartType.AreaStacked,
        ExcelChartType.Pie,
        ExcelChartType.Doughnut,
        ExcelChartType.Scatter,
        ExcelChartType.Radar,
        ExcelChartType.Combo
    };
}

