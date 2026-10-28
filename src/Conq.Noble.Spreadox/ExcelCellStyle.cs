namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Cell style surface (font, fill, borders, alignment, number format).
    /// </summary>
    public sealed class ExcelCellStyle
    {
        public ExcelFont Font { get; } = new ExcelFont();
        public ExcelFill Fill { get; } = new ExcelFill();
        public ExcelBorders Borders { get; } = new ExcelBorders();
        public ExcelAlignment Alignment { get; } = new ExcelAlignment();

        /// <summary>
        /// Excel format code (e.g. <see cref="NumberFormats.Percent"/>). Null = General / type default.
        /// </summary>
        public string? NumberFormat { get; set; }

        public ExcelCellStyle Clone()
        {
            return new ExcelCellStyle
            {
                NumberFormat = NumberFormat,
                // deep-copy nested via assign after construction
            }.CopyFrom(this);
        }

        private ExcelCellStyle CopyFrom(ExcelCellStyle other)
        {
            Font.Name = other.Font.Name;
            Font.Size = other.Font.Size;
            Font.Bold = other.Font.Bold;
            Font.Italic = other.Font.Italic;
            Font.Underline = other.Font.Underline;
            Font.Strikeout = other.Font.Strikeout;
            Font.Color = other.Font.Color;

            Fill.Pattern = other.Fill.Pattern;
            Fill.ForegroundColor = other.Fill.ForegroundColor;
            Fill.BackgroundColor = other.Fill.BackgroundColor;

            Borders.Left.Style = other.Borders.Left.Style;
            Borders.Left.Color = other.Borders.Left.Color;
            Borders.Right.Style = other.Borders.Right.Style;
            Borders.Right.Color = other.Borders.Right.Color;
            Borders.Top.Style = other.Borders.Top.Style;
            Borders.Top.Color = other.Borders.Top.Color;
            Borders.Bottom.Style = other.Borders.Bottom.Style;
            Borders.Bottom.Color = other.Borders.Bottom.Color;

            Alignment.Horizontal = other.Alignment.Horizontal;
            Alignment.Vertical = other.Alignment.Vertical;
            Alignment.WrapText = other.Alignment.WrapText;
            Alignment.Indent = other.Alignment.Indent;
            NumberFormat = other.NumberFormat;
            return this;
        }

        internal bool IsDefault =>
            Font.IsDefault
            && Fill.IsDefault
            && Borders.IsDefault
            && Alignment.IsDefault
            && string.IsNullOrEmpty(NumberFormat);
    }
}
