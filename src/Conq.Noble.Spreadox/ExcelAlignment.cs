namespace Conq.Noble.Spreadox
{
    public sealed class ExcelAlignment
    {
        public ExcelHorizontalAlignment Horizontal { get; set; } = ExcelHorizontalAlignment.General;
        public ExcelVerticalAlignment Vertical { get; set; } = ExcelVerticalAlignment.Bottom;
        public bool WrapText { get; set; }
        public int Indent { get; set; }

        public ExcelAlignment Clone()
        {
            return new ExcelAlignment
            {
                Horizontal = Horizontal,
                Vertical = Vertical,
                WrapText = WrapText,
                Indent = Indent
            };
        }

        internal bool IsDefault =>
            Horizontal == ExcelHorizontalAlignment.General
            && Vertical == ExcelVerticalAlignment.Bottom
            && !WrapText
            && Indent == 0;
    }
}
