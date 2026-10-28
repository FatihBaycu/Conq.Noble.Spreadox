using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Native worksheet chart (OOXML ChartPart). Position uses 1-based cell anchors.
    /// </summary>
    public sealed class ExcelChart
    {
        private readonly List<ExcelChartSeries> _series = new List<ExcelChartSeries>();

        internal ExcelChart(ExcelWorksheet worksheet, ExcelChartType type)
        {
            Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            Type = type;
            FromRow = 2;
            FromColumn = 5;
            ToRow = 15;
            ToColumn = 12;
            if (type == ExcelChartType.BoxAndWhisker)
            {
                ShowLegend = false;
            }
        }

        public ExcelWorksheet Worksheet { get; }

        public ExcelChartType Type { get; }

        public string? Title { get; set; }

        public bool ShowLegend { get; set; } = true;

        public ExcelLegendPosition LegendPosition { get; set; } = ExcelLegendPosition.Right;

        /// <summary>Default data-label options applied when a series does not enable its own.</summary>
        public ExcelDataLabelOptions DataLabels { get; } = new ExcelDataLabelOptions();

        public ExcelChartAxis CategoryAxis { get; } = new ExcelChartAxis { ShowMajorGridlines = false };

        public ExcelChartAxis ValueAxis { get; } = new ExcelChartAxis();

        public ExcelChartAxis SecondaryValueAxis { get; } = new ExcelChartAxis();

        /// <summary>Hole size percent for doughnut charts (default 50).</summary>
        public int DoughnutHoleSize { get; set; } = 50;

        /// <summary>Box-and-whisker layout flags (ignored for other chart types).</summary>
        public ExcelBoxAndWhiskerOptions BoxAndWhisker { get; } = new ExcelBoxAndWhiskerOptions();

        /// <summary>1-based top-left row of the chart anchor.</summary>
        public int FromRow { get; private set; }

        /// <summary>1-based top-left column of the chart anchor.</summary>
        public int FromColumn { get; private set; }

        /// <summary>1-based bottom-right row of the chart anchor.</summary>
        public int ToRow { get; private set; }

        /// <summary>1-based bottom-right column of the chart anchor.</summary>
        public int ToColumn { get; private set; }

        public IReadOnlyList<ExcelChartSeries> Series =>
            new ReadOnlyCollection<ExcelChartSeries>(_series);

        public ExcelChartSeries AddSeries(
            string valuesRange,
            string? categoriesRange = null,
            string? name = null,
            string? nameRange = null)
        {
            Worksheet.Workbook.ThrowIfDisposed();
            if (valuesRange is null) throw new ArgumentNullException(nameof(valuesRange));
            if (string.IsNullOrWhiteSpace(valuesRange))
            {
                throw new ArgumentException("Values range is required.", nameof(valuesRange));
            }

            // Validate A1 shapes early (single cell or range).
            _ = ExcelRange.Parse(valuesRange.Trim());
            if (!string.IsNullOrWhiteSpace(categoriesRange))
            {
                _ = ExcelRange.Parse(categoriesRange!.Trim());
            }

            if (!string.IsNullOrWhiteSpace(nameRange))
            {
                _ = ExcelRange.Parse(nameRange!.Trim());
            }

            var series = new ExcelChartSeries(this)
            {
                ValuesRange = valuesRange.Trim(),
                CategoriesRange = string.IsNullOrWhiteSpace(categoriesRange) ? null : categoriesRange.Trim(),
                Name = name,
                NameRange = string.IsNullOrWhiteSpace(nameRange) ? null : nameRange.Trim()
            };
            _series.Add(series);
            return series;
        }

        public ExcelChart SetPosition(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            if (fromRow <= 0) throw new ArgumentOutOfRangeException(nameof(fromRow));
            if (fromColumn <= 0) throw new ArgumentOutOfRangeException(nameof(fromColumn));
            if (toRow < fromRow) throw new ArgumentOutOfRangeException(nameof(toRow));
            if (toColumn < fromColumn) throw new ArgumentOutOfRangeException(nameof(toColumn));

            FromRow = fromRow;
            FromColumn = fromColumn;
            ToRow = toRow;
            ToColumn = toColumn;
            return this;
        }

        public ExcelChart SetSize(int columnSpan, int rowSpan, int? fromRow = null, int? fromColumn = null)
        {
            if (columnSpan <= 0) throw new ArgumentOutOfRangeException(nameof(columnSpan));
            if (rowSpan <= 0) throw new ArgumentOutOfRangeException(nameof(rowSpan));

            int startRow = fromRow ?? FromRow;
            int startCol = fromColumn ?? FromColumn;
            return SetPosition(startRow, startCol, startRow + rowSpan - 1, startCol + columnSpan - 1);
        }

        internal void AddLoadedSeries(ExcelChartSeries series) => _series.Add(series);

        internal ExcelChartType EffectiveSeriesType(ExcelChartSeries series)
        {
            if (Type == ExcelChartType.Combo)
            {
                return series.SeriesChartType ?? ExcelChartType.ColumnClustered;
            }

            return Type;
        }
    }
}
