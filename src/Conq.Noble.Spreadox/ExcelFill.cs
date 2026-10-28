namespace Conq.Noble.Spreadox
{
    public sealed class ExcelFill
    {
        public ExcelFillPattern Pattern { get; set; } = ExcelFillPattern.None;
        public ExcelColor? ForegroundColor { get; set; }
        public ExcelColor? BackgroundColor { get; set; }

        public static ExcelFill Solid(ExcelColor color)
        {
            return new ExcelFill
            {
                Pattern = ExcelFillPattern.Solid,
                ForegroundColor = color,
                BackgroundColor = color
            };
        }

        public ExcelFill Clone()
        {
            return new ExcelFill
            {
                Pattern = Pattern,
                ForegroundColor = ForegroundColor,
                BackgroundColor = BackgroundColor
            };
        }

        internal bool IsDefault =>
            Pattern == ExcelFillPattern.None
            && ForegroundColor is null
            && BackgroundColor is null;
    }
}
