namespace Conq.Noble.Spreadox
{
    /// <summary>PDF chart axis (in-memory). Not an OOXML axis.</summary>
    public sealed class PdfChartAxis
    {
        public string? Title { get; set; }

        public double? Minimum { get; set; }

        public double? Maximum { get; set; }

        public bool ShowMajorGridlines { get; set; } = true;

        /// <summary>Degrees; negative tilts labels clockwise from horizontal (default −35 for categories).</summary>
        public double LabelRotation { get; set; }
    }
}
