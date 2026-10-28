using System;

namespace Conq.Noble.Spreadox
{
    /// <summary>Page margins in millimeters.</summary>
    public sealed class PdfMargins
    {
        public PdfMargins(double left, double top, double right, double bottom)
        {
            if (left < 0 || top < 0 || right < 0 || bottom < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(left), "Margins cannot be negative.");
            }

            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        public double Left { get; }

        public double Top { get; }

        public double Right { get; }

        public double Bottom { get; }

        public static PdfMargins Millimeters(double left, double top, double right, double bottom) =>
            new PdfMargins(left, top, right, bottom);

        public static PdfMargins Default { get; } = Millimeters(10, 10, 10, 14);
    }
}
