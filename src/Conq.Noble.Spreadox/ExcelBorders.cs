namespace Conq.Noble.Spreadox
{
    public sealed class ExcelBorderEdge
    {
        public ExcelBorderStyle Style { get; set; } = ExcelBorderStyle.None;
        public ExcelColor? Color { get; set; }

        public ExcelBorderEdge Clone()
        {
            return new ExcelBorderEdge
            {
                Style = Style,
                Color = Color
            };
        }

        internal bool IsDefault => Style == ExcelBorderStyle.None && Color is null;
    }

    public sealed class ExcelBorders
    {
        public ExcelBorderEdge Left { get; } = new ExcelBorderEdge();
        public ExcelBorderEdge Right { get; } = new ExcelBorderEdge();
        public ExcelBorderEdge Top { get; } = new ExcelBorderEdge();
        public ExcelBorderEdge Bottom { get; } = new ExcelBorderEdge();

        public void SetAll(ExcelBorderStyle style, ExcelColor? color = null)
        {
            foreach (ExcelBorderEdge edge in new[] { Left, Right, Top, Bottom })
            {
                edge.Style = style;
                edge.Color = color;
            }
        }

        public ExcelBorders Clone()
        {
            var copy = new ExcelBorders();
            copy.Left.Style = Left.Style;
            copy.Left.Color = Left.Color;
            copy.Right.Style = Right.Style;
            copy.Right.Color = Right.Color;
            copy.Top.Style = Top.Style;
            copy.Top.Color = Top.Color;
            copy.Bottom.Style = Bottom.Style;
            copy.Bottom.Color = Bottom.Color;
            return copy;
        }

        internal bool IsDefault =>
            Left.IsDefault && Right.IsDefault && Top.IsDefault && Bottom.IsDefault;
    }
}
