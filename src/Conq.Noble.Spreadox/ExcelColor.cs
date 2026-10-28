using System;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// RGB color for fonts, fills, and borders (public API; not an Open XML type).
    /// </summary>
    public sealed class ExcelColor : IEquatable<ExcelColor>
    {
        public ExcelColor(byte r, byte g, byte b)
        {
            R = r;
            G = g;
            B = b;
        }

        public byte R { get; }
        public byte G { get; }
        public byte B { get; }

        public static ExcelColor FromRgb(byte r, byte g, byte b) => new ExcelColor(r, g, b);

        public static ExcelColor FromHex(string hex)
        {
            if (string.IsNullOrWhiteSpace(hex))
            {
                throw new ArgumentException("Hex color is required.", nameof(hex));
            }

            hex = hex.Trim();
            if (hex.StartsWith("#", StringComparison.Ordinal))
            {
                hex = hex.Substring(1);
            }

            if (hex.Length == 8)
            {
                // AARRGGBB — ignore alpha for OOXML rgb
                hex = hex.Substring(2);
            }

            if (hex.Length != 6)
            {
                throw new ArgumentException("Hex color must be RRGGBB or AARRGGBB.", nameof(hex));
            }

            return new ExcelColor(
                Convert.ToByte(hex.Substring(0, 2), 16),
                Convert.ToByte(hex.Substring(2, 2), 16),
                Convert.ToByte(hex.Substring(4, 2), 16));
        }

        public string ToRgbHex() => $"{R:X2}{G:X2}{B:X2}";

        public bool Equals(ExcelColor? other) =>
            other != null && R == other.R && G == other.G && B == other.B;

        public override bool Equals(object? obj) => Equals(obj as ExcelColor);

        public override int GetHashCode() => (R << 16) | (G << 8) | B;

        public static readonly ExcelColor Black = new ExcelColor(0, 0, 0);
        public static readonly ExcelColor White = new ExcelColor(255, 255, 255);
        public static readonly ExcelColor Red = new ExcelColor(255, 0, 0);
        public static readonly ExcelColor Green = new ExcelColor(0, 128, 0);
        public static readonly ExcelColor Blue = new ExcelColor(0, 0, 255);
        public static readonly ExcelColor Yellow = new ExcelColor(255, 255, 0);
        public static readonly ExcelColor Gray = new ExcelColor(128, 128, 128);
    }
}
