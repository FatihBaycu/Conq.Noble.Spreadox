using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Conq.Noble.Spreadox
{
    public sealed class PdfSummaryItem
    {
        internal PdfSummaryItem(string metric, string value)
        {
            Metric = metric;
            Value = value;
        }

        public string Metric { get; }

        public string Value { get; }
    }

    /// <summary>Right-hand metric:value list (max 28 rows).</summary>
    public sealed class PdfSummaryList
    {
        public const int MaxRows = 28;

        private readonly List<PdfSummaryItem> _items = new List<PdfSummaryItem>();

        public IReadOnlyList<PdfSummaryItem> Items => new ReadOnlyCollection<PdfSummaryItem>(_items);

        public int Count => _items.Count;

        public PdfSummaryList Add(string metric, string value)
        {
            if (string.IsNullOrWhiteSpace(metric))
            {
                throw new ArgumentException("Metric is required.", nameof(metric));
            }

            if (_items.Count < MaxRows)
            {
                _items.Add(new PdfSummaryItem(metric, value ?? string.Empty));
            }

            return this;
        }
    }

    /// <summary>One chart page: title, plot, and a right-hand metric:value list.</summary>
    public sealed class PdfChartPage
    {
        internal PdfChartPage(string title)
        {
            Title = title;
        }

        public string Title { get; set; }

        public PdfChart Chart { get; } = new PdfChart();

        public PdfSummaryList Summary { get; } = new PdfSummaryList();
    }
}
