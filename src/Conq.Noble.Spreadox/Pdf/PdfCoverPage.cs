using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Conq.Noble.Spreadox
{
    public sealed class PdfCoverFact
    {
        internal PdfCoverFact(string label, string value)
        {
            Label = label;
            Value = value;
        }

        public string Label { get; }

        public string Value { get; }
    }

    /// <summary>First page: logo, company, title, date, chart count, facts (e.g. MultiAxis stats).</summary>
    public sealed class PdfCoverPage
    {
        private readonly List<PdfCoverFact> _facts = new List<PdfCoverFact>();

        public string? Company { get; set; }

        public string? Title { get; set; }

        public DateTime? Date { get; set; }

        public PdfImage? Logo { get; set; }

        /// <summary>When null, the writer uses the number of chart pages.</summary>
        public int? ChartCount { get; set; }

        public string? Provider { get; set; }

        public IReadOnlyList<PdfCoverFact> Facts => new ReadOnlyCollection<PdfCoverFact>(_facts);

        public PdfCoverPage AddFact(string label, string value)
        {
            if (string.IsNullOrWhiteSpace(label))
            {
                throw new ArgumentException("Label is required.", nameof(label));
            }

            _facts.Add(new PdfCoverFact(label, value ?? string.Empty));
            return this;
        }
    }
}
