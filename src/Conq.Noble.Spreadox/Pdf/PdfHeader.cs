namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Drawn on every page above the content area. Supports logo, title, and text.
    /// </summary>
    public sealed class PdfHeader
    {
        public string? Title { get; set; }

        public string? Text { get; set; }

        public PdfImage? Logo { get; set; }

        public string? FontFamily { get; set; }

        public double FontSize { get; set; } = 12;

        public bool Bold { get; set; } = true;

        public ExcelColor? TitleColor { get; set; }
    }
}
