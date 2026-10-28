namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Drawn on every page. <see cref="PageNumberFormat"/> uses <c>{0}</c> for page and <c>{1}</c> for total.
    /// </summary>
    public sealed class PdfFooter
    {
        public string PageNumberFormat { get; set; } = "{0} / {1}";

        public string? Text { get; set; }

        public PdfImage? Logo { get; set; }
    }
}
