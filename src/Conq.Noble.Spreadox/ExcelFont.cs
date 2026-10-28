namespace Conq.Noble.Spreadox
{
    public sealed class ExcelFont
    {
        public string? Name { get; set; }
        public double? Size { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikeout { get; set; }
        public ExcelColor? Color { get; set; }

        public ExcelFont Clone()
        {
            return new ExcelFont
            {
                Name = Name,
                Size = Size,
                Bold = Bold,
                Italic = Italic,
                Underline = Underline,
                Strikeout = Strikeout,
                Color = Color
            };
        }

        internal bool IsDefault =>
            string.IsNullOrEmpty(Name)
            && Size is null
            && !Bold
            && !Italic
            && !Underline
            && !Strikeout
            && Color is null;
    }
}
