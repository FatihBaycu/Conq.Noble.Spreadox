using System;

namespace Conq.Noble.Spreadox
{
    /// <summary>Horizontal reference line on the primary value axis (CL / UCL / LCL).</summary>
    public sealed class PdfConstantLine
    {
        internal PdfConstantLine(string name, double value, ExcelColor color)
        {
            Name = name;
            Value = value;
            Color = color;
        }

        public string Name { get; }

        public double Value { get; }

        public ExcelColor Color { get; }
    }
}
